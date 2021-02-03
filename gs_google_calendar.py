import json
import datetime

from extronlib.system import ProgramLog
from gs_calendar_base import _BaseCalendar, _CalendarItem
import gs_requests
import gs_oauth_tools
from gs_service_accounts import _ServiceAccountBase
import time


class GoogleCalendar(_BaseCalendar):
    def __init__(
            self,
            *a,
            getAccessTokenCallback=None,
            calendarName=None,
            debug=False,
            **k
    ):
        if not callable(getAccessTokenCallback):
            raise TypeError('getAccessTokenCallback must be callable')

        self._getAccessTokenCallback = getAccessTokenCallback
        self.calendarName = calendarName
        self._calendarID = None
        self.calendars = set()  # set to avoid duplicates
        self._baseURL = 'https://www.googleapis.com/calendar/v3/'
        self._debug = debug

        super().__init__(
            *a,
            # debug=debug, #nope
            **k)

        self._GetCalendarID()  # init the self.calendars attribute
        self.session = gs_requests.session()

    def __str__(self):
        return '<GoogleCalendar: RoomName={}, LastUpdated={}>'.format(
            self.calendarName,
            self.LastUpdated,
        )

    def print(self, *a, **k):
        if self._debug:
            print(*a, **k)

    def _DoRequest(self, *a, **k):
        self.print('_DoRequest(', a, k)
        if self._GetCalendarID() is None:
            raise PermissionError('Error resolving calendar ID "{}"'.format(self.calendarName))

        self.session.headers['Authorization'] = 'Bearer {}'.format(self._getAccessTokenCallback())
        self.session.headers['Accept'] = 'application/json'
        for key, val in self.session.headers.items():
            if 'Auth' in key:
                val = val[:15] + '...'

            self.print('header', key, '=', val)

        resp = self.session.request(*a, **k)
        return resp

    def _GetCalendarID(self, nextPageToken=None):
        self.print('_GetCalendarID(nextPageToken=', nextPageToken)

        if self._calendarID is None:
            url = self._baseURL + 'users/me/calendarList'.format(
                self._getAccessTokenCallback(),
            )
            if nextPageToken:
                # This will request the next page of results.
                # This happens when the account has access to many-many calendars
                url += '?pageToken={}'.format(nextPageToken)

            self.print('29 url=', url)
            resp = gs_requests.get(
                url,
                headers={
                    'Authorization': 'Bearer {}'.format(self._getAccessTokenCallback())
                }
            )
            self.print('_GetCalendarID resp=', json.dumps(resp.json(), indent=2))
            for calendar in resp.json().get('items', []):
                calendarName = calendar.get('summary', None)

                self.calendars.add(calendarName)

                if calendarName == self.calendarName:
                    self._calendarID = calendar.get('id')
                    self.print('calendar ID found "{}"'.format(self._calendarID))
                    break

            npToken = resp.json().get('nextPageToken', None)
            self.print('npToken=', npToken)
            while self._calendarID is None and npToken is not None:
                self.print('len(items)=', len(resp.json().get('items')))
                return self._GetCalendarID(nextPageToken=npToken)

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

        startDT = startDT or datetime.datetime.now() - datetime.timedelta(days=1)
        endDT = endDT or datetime.datetime.now() + datetime.timedelta(days=7)

        startStr = datetime.datetime.utcfromtimestamp(startDT.timestamp()).isoformat() + "-0000"
        endStr = datetime.datetime.utcfromtimestamp(endDT.timestamp()).isoformat() + "-0000"
        self.print('startStr=', startStr)
        self.print('endStr=', endStr)

        url = self._baseURL + 'calendars/{}/events?timeMax={}&timeMin={}&singleEvents=True'.format(
            self._GetCalendarID(),
            endStr,
            startStr
        )
        resp = self._DoRequest(
            method='get',
            url=url
        )
        self._NewConnectionStatus('Connected' if resp.ok else 'Disconnected')
        self.print('resp=', resp.text)

        theseCalendarItems = []
        for item in resp.json().get('items', []):
            self.print('item=', json.dumps(item, indent=2, sort_keys=True))

            start = fromisoformat(item['start']['dateTime'])
            self.print('95 start=', start)

            end = fromisoformat(item['end']['dateTime'])

            hasAttachments = 'attachments' in item.keys()

            event = _CalendarItem(
                startDT=datetime.datetime.fromtimestamp(start.timestamp()),
                endDT=datetime.datetime.fromtimestamp(end.timestamp()),
                data={
                    'ItemId': item.get('id'),
                    'Subject': item.get('summary'),
                    'OrganizerName': item['creator']['email'],
                    'HasAttachments': hasAttachments,
                    'attachments': item.get('attachments', []),
                },
                parentCalendar=self,
            )
            theseCalendarItems.append(event)

        self.RegisterCalendarItems(
            calItems=theseCalendarItems,
            startDT=startDT,
            endDT=endDT,

        )

    def CreateCalendarEvent(self, subject, body, startDT, endDT):
        timezone = time.tzname[-1] if len(time.tzname) > 1 else time.tzname[0]
        self.print('timezone=', timezone)

        data = {
            "kind": "calendar#event",
            "summary": subject,  # meeting subject
            "description": body,  # meeting body
            "start": {
                # "dateTime": startDT.astimezone(datetime.timezone.utc).isoformat(),# doesnt work on python 3.5
                "dateTime": datetime.datetime.utcfromtimestamp(startDT.timestamp()).isoformat() + '+00:00',
            },
            "end": {
                # "dateTime": endDT.astimezone(datetime.timezone.utc).isoformat(),# doesnt work on python 3.5
                "dateTime": datetime.datetime.utcfromtimestamp(endDT.timestamp()).isoformat() + '+00:00',
            },
        }
        self.print('data=', data)

        resp = self._DoRequest(
            method='POST',
            url='https://www.googleapis.com/calendar/v3/calendars/{calendarID}/events'.format(
                calendarID=self._calendarID,
            ),
            json=data,
        )
        self.print('resp=', resp.text)

        if resp.ok:
            # save the calendar item into memory
            item = resp.json()
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

            self.RegisterCalendarItems(
                calItems=[event],
                startDT=start,
                endDT=end,

            )

    def ChangeEventTime(self, calItem, newStartDT=None, newEndDT=None):
        self.print('ChangeEventTime(', calItem, 'newStartDT=', newStartDT, ', newEndDT=', newEndDT)
        # url = 'http://192.168.68.105'
        url = 'https://www.googleapis.com/calendar/v3/calendars/{calendarId}/events/{eventId}'.format(
            calendarId=self._calendarID,
            eventId=calItem.Get('ItemId')
        )

        data = {
        }

        if newStartDT:
            data['start'] = {
                # "dateTime": newStartDT.astimezone(datetime.timezone.utc).isoformat(),# doesnt work on python 3.5
                "dateTime": datetime.datetime.utcfromtimestamp(newStartDT.timestamp()).isoformat() + '+00:00',
            }

        if newEndDT:
            data['end'] = {
                # "dateTime": newEndDT.astimezone(datetime.timezone.utc).isoformat(),# doesnt work on python 3.5
                "dateTime": datetime.datetime.utcfromtimestamp(newEndDT.timestamp()).isoformat() + '+00:00',
            }

        resp = self._DoRequest(
            method='PATCH',
            url=url,
            json=data,
        )
        self.print('resp=', resp.text)

        if resp.ok:
            # save the calendar item into memory
            item = resp.json()
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

            self.RegisterCalendarItems(
                calItems=[event],
                startDT=start,
                endDT=end,

            )

    def GetAttachments(self, item):
        ret = []
        for d in item.Get('attachments'):
            ret.append(
                _Attachment(
                    AttachmentId=d['fileId'],
                    name=d['title'],
                    parentExchange=self,
                    **d,
                )
            )
        return ret


class _Attachment:
    def __init__(self, AttachmentId, name, parentExchange, **kwargs):
        print('_Attachment(', AttachmentId, parentExchange)
        self.Filename = name
        self.ID = AttachmentId
        self._parentExchange = parentExchange
        self._content = None
        self._kwargs = kwargs

    def Read(self):
        if self._content is None:
            # resp = self._parentExchange.session.get(self._kwargs['fileUrl'])
            resp = self._parentExchange.session.get(
                'https://www.googleapis.com/drive/v3/files/{}'.format(self._kwargs['fileId']))
            print('resp=', resp)
            self._content = resp.content.encode()

        return self._content

    @property
    def Size(self):
        # return size of content in Bytes
        # In theory you could request the size of the attachment from EWS API, or even the hash or changekey
        # but according to this microsoft forum, it is not possible (or at least it does not work as intended)
        # https://social.technet.microsoft.com/Forums/office/en-US/143ab86c-903a-49da-9603-03e65cbd8180/ews-how-to-get-attachment-mime-info-not-content
        return len(self.Read())

    @property
    def Name(self):
        if self.Filename is None:
            self._Update(getContent=False)
        return self.Filename

    def __str__(self):
        return '<Attachment: Name={}>'.format(self.Name)


def fromisoformat(date_string, returnOffsetAware=False):
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

    ret = datetime.datetime(*(date_components + time_components))
    if returnOffsetAware is False:
        ret = datetime.datetime.fromtimestamp(ret.timestamp())
    return ret


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


class ServiceAccount(_ServiceAccountBase):
    def __init__(self, googleJSONpath, oauthID, authManager):
        self.googleJSONpath = googleJSONpath
        self.oauthID = oauthID
        self.authManager = authManager

    def __str__(self):
        return '<Google ServiceAccount: googleJSONpath={}, oauthID={}, authManager={}>'.format(
            self.googleJSONpath,
            self.oauthID[:10] + '...',
            self.authManager,
        )

    @classmethod
    def Dumper(cls, sa):
        return json.dumps({
            'googleJSONpath': sa.googleJSONpath,
            'oauthID': sa.oauthID,
            'authManager': 'devices.authManager',  # todo, generalize this
        }, indent=2, sort_keys=True)

    @classmethod
    def Loader(cls, strng):
        d = json.loads(strng)
        authManager = d.pop('authManager', None)
        if authManager == 'devices.authManager':
            import devices
            authManager = devices.authManager

        print('487 authManager=', authManager)
        print('488 d=', d)

        ret = cls(
            googleJSONpath=d['googleJSONpath'],
            oauthID=d['oauthID'],
            authManager=authManager
        )
        print('Google.ServiceAccount.Loader return', ret)
        return ret

    def GetStatus(self):
        try:
            user = self.authManager.GetUserByID(self.oauthID)
            if user:
                token = user.GetAccessToken()
                if token:
                    return 'Authorized'
                else:
                    return 'Unable to get token'
            else:
                return 'User not found'
        except Exception as e:
            return 'Error 464: {}'.format(e)

    def GetType(self):
        return 'Google'

    def GetRoomInterface(self, roomName, **kwargs):

        user = self.authManager.GetUserByID(self.oauthID)
        if user is None:
            # ProgramLog(
            #     'Google ServiceAccount roomName="{}" kwargs="{}". '
            #     'No User with ID="{}"'.format(
            #         roomName,
            #         kwargs,
            #         self.oauthID
            #     ))
            return

        google = GoogleCalendar(
            getAccessTokenCallback=user.GetAccessToken,
            calendarName=roomName,
            **kwargs,
        )
        return google

    @property
    def calendars(self):
        intf = self.GetRoomInterface(None)
        return intf.calendars


if __name__ == '__main__':
    from gs_oauth_tools import AuthManager
    import webbrowser
    from extronlib import File

    MY_ID = '3888'

    authManager = AuthManager(googleJSONpath='google.json')  # , debug=True)
    user = authManager.GetUserByID(MY_ID)

    if not user:
        d = authManager.CreateNewUser(MY_ID, 'Google')
        try:
            webbrowser.open(d.get('verification_uri'))
        except:
            pass
        print('d=', d)

        while not user:
            user = authManager.GetUserByID(MY_ID)
            time.sleep(1)

        print('user=', user)

    google = GoogleCalendar(
        calendarName='Room Agent Test 34',
        getAccessTokenCallback=user.GetAccessToken,
        debug=True,
        persistentStorage='storage.json',
    )

    google.NewCalendarItem = lambda _, event: print('NewCalendarItem', event)
    google.CalendarItemChanged = lambda _, event: print('CalendarItemChanged', event)
    google.CalendarItemDeleted = lambda _, event: print('CalendarItemDeleted', event)

    # google.CreateCalendarEvent(
    #     subject='Test at {}'.format(time.asctime()),
    #     body='Test Body',
    #     startDT=datetime.datetime.now(),
    #     endDT=datetime.datetime.now() + datetime.timedelta(minutes=15),
    # )

    # while True:
    #     google.UpdateCalendar(
    #         startDT=datetime.datetime.now(),
    #         endDT=datetime.datetime.now() + datetime.timedelta(days=7),
    #     )
    #     time.sleep(10)

    # time.sleep(2)
    google.UpdateCalendar(
        # startDT=datetime.datetime.now().replace(hour=0, minute=0, microsecond=0),
        # endDT=datetime.datetime.now() + datetime.timedelta(days=1),
    )

    nowEvents = google.GetNowCalItems()
    print('nowEvents=', nowEvents)
    # print('all Events=', google.GetAllEvents())
    for event in nowEvents:
        print('event=', event)
        # google.ChangeEventTime(
        #     event,
        #     newStartDT=event.Get('Start') - datetime.timedelta(minutes=15),
        #     newEndDT=event.Get('End') + datetime.timedelta(minutes=15),
        # )

        for attachment in event.Attachments:
            print('attachment=', attachment)
            with File(attachment.Name, mode='wb') as file:
                file.write(attachment.Read())
