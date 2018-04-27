# -*- coding: utf-8 -*-

import inspect
import re
import types
from datetime import datetime
from itertools import chain, groupby
from xlwt import Formula

from trac.core import Component, implements
from trac.env import Environment
from trac.mimeview.api import Context, IContentConverter
from trac.resource import Resource, get_resource_url
from trac.ticket.api import TicketSystem
from trac.ticket.model import Ticket
from trac.ticket.query import Query
from trac.ticket.web_ui import TicketModule
from trac.util import Ranges
from trac.util.text import empty, unicode_urlencode
from trac.web.api import IRequestFilter, RequestDone
from trac.web.chrome import Chrome, add_link
try:
    from trac.util.datefmt import from_utimestamp
except ImportError:
    from datetime import timedelta
    from trac.util.datefmt import utc
    _epoc = datetime(1970, 1, 1, tzinfo=utc)
    from_utimestamp = lambda ts: _epoc + timedelta(seconds=ts or 0)

from tracexceldownload.api import (get_excel_format, get_excel_mimetype,
                                   get_workbook_writer, get_literal)
from tracexceldownload.translation import _, dgettext, dngettext


if hasattr(Environment, 'get_read_db'):
    _get_db = lambda env: env.get_read_db()
else:
    _get_db = lambda env: env.get_db_cnx()


def _tkt_id_conditions(column, tkt_ids):
    ranges = Ranges()
    ranges.appendrange(','.join(map(str, sorted(tkt_ids))))
    condition = []
    tkt_ids = []
    for a, b in ranges.pairs:
        if a == b:
            tkt_ids.append(a)
        elif a + 1 == b:
            tkt_ids.extend((a, b))
        else:
            condition.append('%s BETWEEN %d AND %d' % (column, a, b))
    if tkt_ids:
        condition.append('%s IN (%s)' % (column, ','.join(map(str, tkt_ids))))
    return ' OR '.join(condition)


class BulkFetchTicket(Ticket):

    @classmethod
    def select(cls, env, tkt_ids):
        if not tkt_ids:
            return {}

        db = _get_db(env)
        fields = TicketSystem(env).get_ticket_fields()
        std_fields = [f['name'] for f in fields if not f.get('custom')]
        time_fields = [f['name'] for f in fields if f['type'] == 'time']
        custom_fields = set(f['name'] for f in fields if f.get('custom'))
        cursor = db.cursor()
        tickets = {}

        cursor.execute('SELECT %s,id FROM ticket WHERE %s' %
                       (','.join(std_fields),
                        _tkt_id_conditions('id', tkt_ids)))
        for row in cursor:
            id = row[-1]
            values = {}
            for idx, field in enumerate(std_fields):
                value = row[idx]
                if field in time_fields:
                    value = from_utimestamp(value)
                elif value is None:
                    value = empty
                values[field] = value
            tickets[id] = (values, [])  # values, changelog

        cursor.execute('SELECT ticket,name,value FROM ticket_custom '
                       'WHERE %s ORDER BY ticket' %
                       _tkt_id_conditions('ticket', tkt_ids))
        for id, rows in groupby(cursor, lambda row: row[0]):
            if id not in tickets:
                continue
            values = {}
            for id, name, value in rows:
                if name in custom_fields:
                    if value is None:
                        value = empty
                    values[name] = value
            tickets[id][0].update(values)

        cursor.execute('SELECT ticket,time,author,field,oldvalue,newvalue '
                       'FROM ticket_change WHERE %s ORDER BY ticket,time' %
                       _tkt_id_conditions('ticket', tkt_ids))
        for id, rows in groupby(cursor, lambda row: row[0]):
            if id not in tickets:
                continue
            tickets[id][1].extend(
                    (from_utimestamp(t), author, field, oldvalue or '',
                     newvalue or '', 1)
                    for id, t, author, field, oldvalue, newvalue in rows)

        return dict((id, cls(env, id, values=values, changelog=changelog,
                             fields=fields, time_fields=time_fields))
                    for id, (values, changelog) in tickets.iteritems())

    def __init__(self, env, tkt_id=None, db=None, version=None, values=None,
                 changelog=None, fields=None, time_fields=None):
        self.env = env
        if tkt_id is not None:
            tkt_id = int(tkt_id)
        self.fields = fields
        self.time_fields = time_fields
        self.id = tkt_id
        self.version = version
        self._values = values
        self.values = values.copy()
        self._changelog = changelog
        self._old = {}

    @property
    def resource(self):
        return Resource('ticket', self.id, self.version)

    def _fetch_ticket(self, tkt_id, db=None):
        self.values = self._values.copy()

    def get_changelog(self, when=None, db=None):
        return self._changelog[:]


class ExcelTicketModule(Component):

    implements(IContentConverter)

    def get_supported_conversions(self):
        format = get_excel_format(self.env)
        mimetype = get_excel_mimetype(format)
        yield ('excel', _("Excel"), format,
               'trac.ticket.Query', mimetype, 8)
        yield ('excel-history', _("Excel including history"), format,
               'trac.ticket.Query', mimetype, 8)
        yield ('excel-history', _("Excel including history"), format,
               'trac.ticket.Ticket', mimetype, 8)

    def convert_content(self, req, mimetype, content, key):
        if key == 'excel':
            return self._convert_query(req, content)
        if key == 'excel-history':
            kwargs = {}
            if isinstance(content, Ticket):
                content = Query.from_string(self.env, 'id=%d' % content.id)
                kwargs['sheet_query'] = False
                kwargs['sheet_history'] = True
            else:
                kwargs['sheet_query'] = True
                kwargs['sheet_history'] = True
            return self._convert_query(req, content, **kwargs)

    def _convert_query(self, req, query, sheet_query=True,
                       sheet_history=False):
        book = get_workbook_writer(self.env, req)

        # no paginator
        query.max = 0
        query.has_more_pages = False
        query.offset = 0
        db = _get_db(self.env)

        # extract all fields except custom fields
        custom_fields = [f['name'] for f in query.fields if f.get('custom')]
        cols = ['id']
        cols.extend(f['name'] for f in query.fields
                              if f['name'] not in custom_fields)
        cols.extend(name for name in ('time', 'changetime')
                         if name not in cols)
        query.cols = cols

        # prevent "SELECT COUNT(*)" query
        saved_count_prop = query._count
        try:
            query._count = types.MethodType(lambda self, sql, args, db=None: 0,
                                            query, query.__class__)
            if 'db' in inspect.getargspec(query.execute)[0]:
                tickets = query.execute(req, db)
            else:
                tickets = query.execute(req)
            query.num_items = len(tickets)
        finally:
            query._count = saved_count_prop

        # add custom fields to avoid error to join many tables
        self._fill_custom_fields(tickets, query.fields, custom_fields, db)

        context = Context.from_request(req, 'query', absurls=True)
        cols.extend([name for name in custom_fields if name not in cols])
        data = query.template_data(context, tickets)

        if sheet_query:
            self._create_sheet_query(req, context, data, book)
        if sheet_history:
            self._create_sheet_history(req, context, data, book)
        return book.dumps(), book.mimetype

    def _fill_custom_fields(self, tickets, fields, custom_fields, db):
        if not tickets or not custom_fields:
            return
        fields = dict((f['name'], f) for f in fields)
        tickets = dict((int(ticket['id']), ticket) for ticket in tickets)
        query = "SELECT ticket,name,value " \
                "FROM ticket_custom WHERE %s ORDER BY ticket" % \
                _tkt_id_conditions('ticket', tickets)

        cursor = db.cursor()
        cursor.execute(query)
        for id, name, value in cursor:
            if id not in tickets:
                continue
            f = fields.get(name)
            if f and f['type'] == 'checkbox':
                try:
                    value = bool(int(value))
                except (TypeError, ValueError):
                    value = False
            tickets[id][name] = value

    def _create_sheet_query(self, req, context, data, book):
        def write_headers(writer, query):
            writer.write_row([(
                u'%s (%s)' % (dgettext('messages', 'Custom Query'),
                              dngettext('messages', '%(num)s match',
                                        '%(num)s matches', query.num_items)),
                'header', -1, -1)])

        query = data['query']
        groups = data['groups']
        fields = data['fields']
        headers = data['headers']

        sheet_count = 1
        sheet_name = dgettext("messages", "Custom Query")
        writer = book.create_sheet(sheet_name)
        write_headers(writer, query)

        for groupname, results in groups:
            results = [result for result in results
                              if 'TICKET_VIEW' in req.perm(
                                 context('ticket', result['id']).resource)]
            if not results:
                continue

            if writer.row_idx + len(results) + 3 > writer.MAX_ROWS:
                sheet_count += 1
                writer = book.create_sheet('%s (%d)' % (sheet_name,
                                                        sheet_count))
                write_headers(writer, query)

            if groupname:
                writer.move_row()
                cell = fields[query.group]['label'] + ' '
                if query.group in ('owner', 'reporter'):
                    cell += Chrome(self.env).format_author(req, groupname)
                else:
                    cell += groupname
                cell += ' (%s)' % dngettext('messages', '%(num)s match',
                                            '%(num)s matches', len(results))
                writer.write_row([(cell, 'header2', -1, -1)])

            writer.write_row((header['label'], 'thead', None, None)
                             for idx, header in enumerate(headers))

            for result in results:
                ticket_context = context('ticket', result['id'])
                cells = []
                for idx, header in enumerate(headers):
                    name = header['name']
                    value, style, width, line = self._get_cell_data(
                        name, result.get(name), req, ticket_context, writer)
                    cells.append((value, style, width, line))
                writer.write_row(cells)

        writer.set_col_widths()

    def _create_sheet_history(self, req, context, data, book):
        def write_headers(writer, headers):
            writer.write_row((header['label'], 'thead', None, None)
                             for idx, header in enumerate(headers))

        groups = data['groups']
        headers = [header for header in data['headers']
                   if header['name'] not in ('id', 'time', 'changetime')]
        headers[0:0] = [
            {'name': 'id', 'label': dgettext("messages", "Ticket")},
            {'name': 'time', 'label': dgettext("messages", "Time")},
            {'name': 'author', 'label': dgettext("messages", "Author")},
            {'name': 'comment', 'label': dgettext("messages", "Comment")},
        ]

        sheet_name = dgettext("messages", "Change History")
        sheet_count = 1
        writer = book.create_sheet(sheet_name)
        write_headers(writer, headers)

        tkt_ids = [result['id']
                   for result in chain(*[results for groupname, results
                                                 in groups])]
        tickets = BulkFetchTicket.select(self.env, tkt_ids)

        mod = TicketModule(self.env)
        for result in chain(*[results for groupname, results in groups]):
            id = result['id']
            ticket = tickets[id]
            ticket_context = context('ticket', id)
            if 'TICKET_VIEW' not in req.perm(ticket_context.resource):
                continue
            values = ticket.values.copy()
            changes = []

            for change in mod.grouped_changelog_entries(ticket, None):
                if change['permanent']:
                    changes.append(change)
            for change in reversed(changes):
                change['values'] = values
                values = values.copy()
                for name, field in change['fields'].iteritems():
                    if name in values:
                        values[name] = field['old']
            changes[0:0] = [{'date': ticket.time_created, 'fields': {},
                             'values': values, 'cnum': None,
                             'comment': '', 'author': ticket['reporter']}]

            if writer.row_idx + len(changes) >= writer.MAX_ROWS:
                sheet_count += 1
                writer = book.create_sheet('%s (%d)' % (sheet_name,
                                                        sheet_count))
                write_headers(writer, headers)

            for change in changes:
                cells = []
                for idx, header in enumerate(headers):
                    name = header['name']
                    if name == 'id':
                        value = id
                    elif name == 'time':
                        value = change.get('date', '')
                    elif name == 'comment':
                        value = change.get('comment', '')
                    elif name == 'author':
                        value = change.get('author', '')
                        value = Chrome(self.env).format_author(req, value)
                    else:
                        value = change['values'].get(name, '')
                    value, style, width, line = \
                            self._get_cell_data(name, value, req,
                                                ticket_context, writer)
                    if name in change['fields']:
                        style = '%s:change' % style
                    cells.append((value, style, width, line))
                writer.write_row(cells)

        writer.set_col_widths()

    def _get_cell_data(self, name, value, req, context, writer):

        if name == 'tt_spent':
            if not value:
                value = 0
                return value, name, None, None
            width = len(value)
            try:
                value = float(value)
            except ValueError:
                pass  # leave value as is
            return value, name, width, None

        if name == 'tt_estimated':
            if not value:
                value = 0
                return value, name, None, None
            width = len(value)
            try:
                value = float(value)
            except ValueError:
                pass  # leave value as is
            return value, name, width, None

        if name == 'tt_remaining':
            if not value:
                value = 0
                return value, name, None, None
            width = len(value)
            try:
                value = float(value)
            except ValueError:
                pass  # leave value as is
            return value, name, width, None

        if name == 'br_planned':
            if not value:
                value = 0
                return value, name, None, None
            width = len(value)
            try:
                value = float(value)
            except ValueError:
                pass  # leave value as is
            return value, name, width, None

        if name == 'children':
            # this never has a value for some reason
            pass

        if name == 'parent':
            if not value:
                value = ''
                return value, name, None, None
            values = re.findall('[\d]+', value)
            parents = ''
            if len(values) == 1:
                value = values[0]
                url = self.env.abs_href.ticket(value)
                value = '#%d' % int(value)
                value = Formula('HYPERLINK("%s",%s)' % (url, get_literal(value)))
                width = len(values[0])
                return value, name, width, 1
            for value in values:
                # hyperlinks aren't working with multiple parents, set as string
                url = self.env.abs_href.ticket(value)
                value = '#%d' % int(value)
                parents += value + ' '
            width = len(values) * 3
            return parents, name, width, 1

        if name == 'id':
            url = self.env.abs_href.ticket(value)
            value = '#%d' % value
            width = len(value)
            value = Formula('HYPERLINK("%s",%s)' % (url, get_literal(value)))
            return value, 'id', width, 1

        if isinstance(value, datetime):
            return value, '[datetime]', None, None

        if value and name in ('reporter', 'owner'):
            value = Chrome(self.env).format_author(req, value)
            return value, name, None, None

        if name == 'cc':
            value = Chrome(self.env).format_emails(context, value)
            return value, name, None, None

        if name == 'milestone':
            if value:
                url = self.env.abs_href.milestone(value)
                width, line = writer.get_metrics(value)
                return value, name, width, line
            else:
                return '', name, None, None

        return value, name, None, None


class ExcelReportModule(Component):

    implements(IRequestFilter)

    _PATH_INFO_MATCH = re.compile(r'/report/[0-9]+').match

    def pre_process_request(self, req, handler):
        if self._PATH_INFO_MATCH(req.path_info) \
                and req.args.get('format') in ('xlsx', 'xls') \
                and handler.__class__.__name__ == 'ReportModule':
            req.args['max'] = 0
        return handler

    def post_process_request(self, req, template, data, content_type):
        if template == 'report_view.html' and req.args.get('id'):
            format = req.args.getfirst('format')
            if format in ('xlsx', 'xls'):
                resource = Resource('report', req.args['id'])
                data['context'] = Context.from_request(req, resource,
                                                       absurls=True)
                self._convert_report(format, req, data)
            elif not format:
                self._add_alternate_links(req)
        return template, data, content_type

    def _convert_report(self, format, req, data):
        book = get_workbook_writer(self.env, req)
        writer = book.create_sheet(dgettext('messages', 'Report'))

        writer.write_row([(
            '%s (%s)' % (data['title'],
                         dngettext('messages', '%(num)s match',
                                   '%(num)s matches', data['numrows'])),
            'header', -1, -1)])

        for value_for_group, row_group in data['row_groups']:
            writer.move_row()

            if value_for_group and len(row_group):
                writer.write_row([(
                    '%s (%s)' % (value_for_group,
                                 dngettext('messages', '%(num)s match',
                                           '%(num)s matches', len(row_group))),
                    'header2', -1, -1)])
            for header_group in data['header_groups']:
                writer.write_row([
                    (header['title'], 'thead', None, None)
                    for header in header_group
                    if not header['hidden']])

            for row in row_group:
                for cell_group in row['cell_groups']:
                    cells = []
                    for cell in cell_group:
                        cell_header = cell['header']
                        if cell_header['hidden']:
                            continue
                        col = cell_header['col'].strip('_').lower()
                        value, style, width, line = \
                            self._get_cell_data(req, col, cell, row, writer)
                        cells.append((value, style, width, line))
                    writer.write_row(cells)

        writer.set_col_widths()

        content = book.dumps()
        req.send_response(200)
        req.send_header('Content-Type', book.mimetype)
        req.send_header('Content-Length', len(content))
        req.send_header('Content-Disposition',
                        'filename=report_%s.%s' % (req.args['id'], format))
        req.end_headers()
        req.write(content)
        raise RequestDone

    def _get_cell_data(self, req, col, cell, row, writer):
        value = cell['value']
        
        if col == 'tt_spent':
            if not value:
                value = 0
                return value, col, None, None
            width = len(value)
            value = float(re.findall('[\d\.]+', cell['value'])[0])
            return value, col, width, None

        if col == 'tt_estimated':
            if not value:
                value = 0
                return value, col, None, None
            width = len(value)
            value = float(re.findall('[\d\.]+', cell['value'])[0])
            return value, col, width, None

        if col == 'tt_remaining':
            if not value:
                value = 0
                return value, col, None, None
            width = len(value)
            value = float(re.findall('[\d\.]+', cell['value'])[0])
            return value, col, width, None

        if col == 'br_planned':
            if not value:
                value = 0
                return value, col, None, None
            width = len(value)
            try:
                value = float(value)
            except ValueError:
                pass  # leave value as is
            return value, col, width, None

        if col == 'parent':
            if not value:
                value = ''
                return value, col, None, None
            value = int(float(re.findall('[\d]+', value)[0]))
            url = self.env.abs_href.ticket(value)
            value = '#%d' % value
            width = len(value)
            value = Formula('HYPERLINK("%s",%s)' % (url, get_literal(value)))
            return value, col, width, 1

        if col == 'report':
            url = self.env.abs_href.report(value)
            width, line = writer.get_metrics(value)
            return value, col, width, line

        if col in ('ticket', 'id'):
            id_value = cell['value']
            value = '#%s' % id_value
            url = get_resource_url(self.env, row['resource'], self.env.abs_href)
            width = len(value)
            return id_value, 'id', width, 1

        if col == 'milestone':
            url = self.env.abs_href.milestone(value)
            width, line = writer.get_metrics(value)
            return value, col, width, line

        if col == 'time':
            if isinstance(value, basestring) and value.isdigit():
                value = from_utimestamp(long(value))
                return value, '[time]', None, None
        elif col in ('date', 'created', 'modified'):
            if isinstance(value, basestring) and value.isdigit():
                value = from_utimestamp(long(value))
                return value, '[date]', None, None
        elif col == 'datetime':
            if isinstance(value, basestring) and value.isdigit():
                value = from_utimestamp(long(value))
                return value, '[datetime]', None, None

        width, line = writer.get_metrics(value)
        return value, col, width, line

    def _add_alternate_links(self, req):
        params = {}
        for arg in req.args.keys():
            if not arg.isupper():
                continue
            params[arg] = req.args.get(arg)
        if 'USER' not in params:
            params['USER'] = req.authname
        if 'sort' in req.args:
            params['sort'] = req.args['sort']
        if 'asc' in req.args:
            params['asc'] = req.args['asc']
        href = ''
        if params:
            href = '&' + unicode_urlencode(params)
        format = get_excel_format(self.env)
        mimetype = get_excel_mimetype(format)
        add_link(req, 'alternate', '?format=' + format + href,
                 _("Excel"), mimetype)
