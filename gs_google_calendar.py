import json
import datetime
from gs_calendar_base import _BaseCalendar, _CalendarItem
import gs_requests


class GoogleCalendar(_BaseCalendar):
    def __init__(self, *a, getAccessTokenCallback=None, calendarName=None, debug=False, **k):
        if not callable(getAccessTokenCallback):
            raise TypeError('getAccessTokenCallback must be callable')

        self._getAccessTokenCallback = getAccessTokenCallback
        self._calendarName = calendarName
        self._calendarID = None
        self.calendars = set()  # set to avoid duplicates
        self._baseURL = 'https://www.googleapis.com/calendar/v3/'
        self._debug = debug

        super().__init__(*a, **k)

        self._getCalendarID()  # init the self.calendars attribute

    def print(self, *a, **k):
        if self._debug:
            print(*a, **k)

    def _DoRequest(self, *a, **k):
        if self._getCalendarID() is None:
            raise PermissionError('No calendar ID')

        return gs_requests.request(*a, **k)

    def _getCalendarID(self):
        if self._calendarID is None:
            url = self._baseURL + 'users/me/calendarList?access_token={0}'.format(
                self._getAccessTokenCallback(),
            )
            self.print('29 url=', url)
            resp = gs_requests.get(url)
            self._NewConnectionStatus('Connected' if resp.ok else 'Disconnected')
            self.print('_getCalendarID resp=', json.dumps(resp.json(), indent=2))
            for calendar in resp.json().get('items', []):
                calendarName = calendar.get('summary', None)

                self.calendars.add(calendarName)

                if calendarName == self._calendarName:
                    self._calendarID = calendar.get('id')
                    self.print('New calendar ID found "{}"'.format(self._calendarID))
                    break

        return self._calendarID

    def UpdateCalendar(self, calendar=None, startDT=None, endDT=None):
        '''
        Subclasses should override this

        :param calendar: a particular calendar ( None means use the default calendar)
        :param startDT: only search for events after this date
        :param endDT: only search for events before this date
        :return:
        '''
        self.print('UpdateCalendar(', calendar, startDT, endDT)

        startDT = startDT or datetime.datetime.utcnow()
        endDT = endDT or datetime.datetime.utcnow() + datetime.timedelta(days=7)

        startStr = startDT.replace(microsecond=0).isoformat() + "-0000"
        endStr = endDT.replace(microsecond=0).isoformat() + "-0000"
        url = self._baseURL + 'calendars/{}/events?access_token={}&timeMax={}&timeMin={}&singleEvents=True'.format(
            self._getCalendarID(),
            self._getAccessTokenCallback(),
            endStr,
            startStr
        )
        resp = self._DoRequest(
            method='get',
            url=url
        )
        self.print('resp=', resp.text)
        self._NewConnectionStatus('Connected' if resp.ok else 'Disconnected')

        theseCalendarItems = []
        for item in resp.json().get('items', []):
            self.print('item=', json.dumps(item, indent=2, sort_keys=True))

            start = fromisoformat(item['start']['dateTime'])
            # start is offset-aware

            end = fromisoformat(item['end']['dateTime'])

            event = _CalendarItem(
                startDT=datetime.datetime.fromtimestamp(start.timestamp()),
                endDT=datetime.datetime.fromtimestamp(end.timestamp()),
                data={
                    'ItemId': item.get('id'),
                    'Subject': item.get('summary'),
                    'OrganizerName': item['creator']['email'],
                    'HasAttachment': False,
                },
                parentCalendar=self,
            )
            theseCalendarItems.append(event)

        self.RegisterCalendarItems(
            calItems=theseCalendarItems,
            startDT=startDT,
            endDT=endDT,

        )


def fromisoformat(date_string):
    # apparently GS python 3.5 does not support this method.
    # Copied these from python 3.8 datetime source code

    """Construct a datetime from the output of datetime.isoformat()."""
    if not isinstance(date_string, str):
        raise TypeError('fromisoformat: argument must be str')

    # Split this at the separator
    dstr = date_string[0:10]
    tstr = date_string[11:]

    try:
        date_components = _parse_isoformat_date(dstr)
    except ValueError:
        raise ValueError('Invalid isoformat string: {date_string}'.format(date_string=date_string))

    if tstr:
        try:
            time_components = _parse_isoformat_time(tstr)
        except ValueError:
            raise ValueError('Invalid isoformat string: {date_string}'.format(date_string=date_string))
    else:
        time_components = [0, 0, 0, 0, None]

    return datetime.datetime(*(date_components + time_components))


def _parse_isoformat_time(tstr):
    # Format supported is HH[:MM[:SS[.fff[fff]]]][+HH:MM[:SS[.ffffff]]]
    len_str = len(tstr)
    if len_str < 2:
        raise ValueError('Isoformat time too short')

    # This is equivalent to re.search('[+-]', tstr), but faster
    tz_pos = (tstr.find('-') + 1 or tstr.find('+') + 1)
    timestr = tstr[:tz_pos - 1] if tz_pos > 0 else tstr

    time_comps = _parse_hh_mm_ss_ff(timestr)

    tzi = None
    if tz_pos > 0:
        tzstr = tstr[tz_pos:]

        # Valid time zone strings are:
        # HH:MM               len: 5
        # HH:MM:SS            len: 8
        # HH:MM:SS.ffffff     len: 15

        if len(tzstr) not in (5, 8, 15):
            raise ValueError('Malformed time zone string')

        tz_comps = _parse_hh_mm_ss_ff(tzstr)
        if all(x == 0 for x in tz_comps):
            tzi = datetime.timezone.utc
        else:
            tzsign = -1 if tstr[tz_pos - 1] == '-' else 1

            td = datetime.timedelta(hours=tz_comps[0], minutes=tz_comps[1],
                                    seconds=tz_comps[2], microseconds=tz_comps[3])

            tzi = datetime.timezone(tzsign * td)

    time_comps.append(tzi)

    return time_comps


def _parse_hh_mm_ss_ff(tstr):
    # Parses things of the form HH[:MM[:SS[.fff[fff]]]]
    len_str = len(tstr)

    time_comps = [0, 0, 0, 0]
    pos = 0
    for comp in range(0, 3):
        if (len_str - pos) < 2:
            raise ValueError('Incomplete time component')

        time_comps[comp] = int(tstr[pos:pos + 2])

        pos += 2
        next_char = tstr[pos:pos + 1]

        if not next_char or comp >= 2:
            break

        if next_char != ':':
            raise ValueError('Invalid time separator: %c' % next_char)

        pos += 1

    if pos < len_str:
        if tstr[pos] != '.':
            raise ValueError('Invalid microsecond component')
        else:
            pos += 1

            len_remainder = len_str - pos
            if len_remainder not in (3, 6):
                raise ValueError('Invalid microsecond component')

            time_comps[3] = int(tstr[pos:])
            if len_remainder == 3:
                time_comps[3] *= 1000

    return time_comps


def _parse_isoformat_date(dtstr):
    # It is assumed that this function will only be called with a
    # string of length exactly 10, and (though this is not used) ASCII-only
    year = int(dtstr[0:4])
    if dtstr[4] != '-':
        raise ValueError('Invalid date separator: %s' % dtstr[4])

    month = int(dtstr[5:7])

    if dtstr[7] != '-':
        raise ValueError('Invalid date separator')

    day = int(dtstr[8:10])

    return [year, month, day]


if __name__ == '__main__':
    from gs_oauth_tools import AuthManager
    import time
    import webbrowser

    MY_ID = '3888'

    authManager = AuthManager(googleJSONpath='google_test_creds.json')
    user = authManager.GetUserByID(MY_ID)

    if not user:
        d = authManager.CreateNewUser(MY_ID, 'Google')
        webbrowser.open(d.get('verification_uri'))
        print('d=', d)

        while not user:
            user = authManager.GetUserByID(MY_ID)
            time.sleep(1)

        print('user=', user)

    google = GoogleCalendar(
        calendarName='Room Agent Test',
        getAccessTokenCallback=user.GetAcessToken,
        debug=True,
    )

    google.NewCalendarItem = lambda _, event: print('NewCalendarItem', event)
    google.CalendarItemChanged = lambda _, event: print('CalendarItemChanged', event)
    google.CalendarItemDeleted = lambda _, event: print('CalendarItemDeleted', event)

    while True:
        google.UpdateCalendar(
            startDT=datetime.datetime.utcnow(),
            endDT=datetime.datetime.utcnow() + datetime.timedelta(days=7),
        )
        time.sleep(10)
