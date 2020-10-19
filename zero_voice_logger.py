import win32con
import win32event
import win32evtlog
import datetime
import time
import os
import threading
import cherrypy
import webbrowser
import collections
import itertools
import json
import shutil


import pprint


server = 'localhost' # get local events
logqry = '*[System[(EventID=4663)]]' # Filter access event
logtype = 'Security' # file access logged in security logs

file_filter = '\\se\\ed7v' # only files from se\ed7v.....wav
proc_filter = 'ED_ZERO.exe' # filter process
#uncomment line below for testing with windows media player
proc_filter = 'wmplayer.exe' # filter process

json_file = './data/zero.json'

user_config_file = 'user_config.json'
user_config = {}

voices = collections.OrderedDict()
voice_data_dict = {
    'last_occur': None,
    'status': None,
    'voice': '',
    'speaker': '',
    'spoken_when': '',
    'comment': '',
    'translation_language': 'EN',
    'translation': '',
    'need_sfx': '',
    'no_speech': '',
    }


def extract_voice_file(xml):
    """extract file name from event xml"""
    p1 = xml.find(file_filter)
    p2 = xml.find('<', p1)
    return xml[p1+4:p2]


def subscribe_and_yield_events(channel, query='*'):
    """subcribe and report events for voice files access by proc_filter"""
    h = win32event.CreateEvent(None, 0, 0, None)
    s = win32evtlog.EvtSubscribe(channel, win32evtlog.EvtSubscribeToFutureEvents, SignalEvent=h, Query=query)

    while True:
        while True:
            events = win32evtlog.EvtNext(s, 1)

            if len(events) == 0:
                break
            for event in events:
                xml = win32evtlog.EvtRender(event, win32evtlog.EvtRenderEventXml)
                if file_filter in xml and proc_filter in xml:
                    yield extract_voice_file(xml)

        while True:
            w = win32event.WaitForSingleObjectEx(h, 100, True)
            if w == win32con.WAIT_OBJECT_0:
                break


def collect_voice(voice):
    """insert new and update old voice data"""
    now = datetime.datetime.now().strftime("%Y%m%d %H:%M:%S")
    voice_data = voices[voice]
    voice_data['status'] = 'new'
    voice_data['last_occur'] = now
    voices[voice] = voice_data


class zeroVoice(object):


    def __init__(self):
        self.voices_page_size = 10
        self.voices_page = 1


    def read_template(self, file):
        with open(file) as f:
            return f.read()
            

    def updatevoices(self, data):
        #convert single data of row into list
        if type(data['voice']) != type([]):
            for key in data:
                temp = []
                temp.append(data[key])
                data[key] = temp
        for key, voice in enumerate(data['voice']):
            voices[voice]['speaker'] = data['speaker'][key]
            voices[voice]['spoken_when'] = data['spoken_when'][key]
            voices[voice]['comment'] = data['comment'][key]
            voices[voice]['translation_language'] = data['translation_language'][key]
            voices[voice]['translation'] = data['translation'][key]
            voices[voice]['status'] = ''
        else:
            now = datetime.datetime.now().strftime("_%Y%m%d_%H%M")
            shutil.copyfile(json_file, json_file + now)
            with open(json_file, 'w') as f:
                json.dump(voices, f, indent=4)


    def sortlastvoices(self, voice):
        if voices[voice]['last_occur'] == None:
            return ""
        return voices[voice]['last_occur']

    
    @cherrypy.expose
    def lastvoices(self):
        """show last voices in webpage"""
        template_file = './public/template_newvoices.html'
        template_entry_file = './public/template_voices_entry.html'
        template_datalist_file = './public/template_datalist.txt'
        
        template = self.read_template(template_file)
        template_entry = self.read_template(template_entry_file)
        template_datalist = self.read_template(template_datalist_file)

        voices_html = ''
        for row, voice in enumerate(reversed(sorted(voices, key=self.sortlastvoices))):
            if voices[voice]['status'] == None:
                break
            data = voices[voice]
            voice_html = template_entry.format(voiceid=row, **data)
            voices_html = voices_html + voice_html
        html = template.format(datalist_html=template_datalist, voices_html=voices_html)
        return html


    @cherrypy.expose
    def voices(self):
        """show all voices for editing"""
        template_file = './public/template_voices.html'
        template_entry_file = './public/template_voices_entry.html'
        template_datalist_file = './public/template_datalist.txt'
        
        template = self.read_template(template_file)
        template_entry = self.read_template(template_entry_file)
        template_datalist = self.read_template(template_datalist_file)

        voices_html = ''
        page_from = (self.voices_page - 1) * self.voices_page_size
        page_to = (self.voices_page - 1) * self.voices_page_size + self.voices_page_size
        for row, voice in enumerate(dict(itertools.islice(voices.items(),
                                                          page_from,
                                                          page_to))):
            data = voices[voice]
            voice_html = template_entry.format(voiceid=row, **data)
            voices_html = voices_html + voice_html
        html = template.format(datalist_html=template_datalist, voices_html=voices_html)
        return html


    @cherrypy.expose
    def voices_next(self):
        self.voices_page = self.voices_page + 1
        raise cherrypy.HTTPRedirect('/voices')


    @cherrypy.expose
    def voices_prev(self):
        self.voices_page = self.voices_page - 1
        if self.voices_page < 1:
            self.voices_page = 1
        raise cherrypy.HTTPRedirect('/voices')


    @cherrypy.expose
    def voice_save(self, **postdata):
        """save posted voice data to json file, update voices dict"""
        if postdata:
            pprint.pprint(postdata)
            voices[postdata['voice']] = {**voices[postdata['voice']], **postdata}
            filename = './data/{voice}.json'.format(voice=postdata['voice'][:-4])
            with open(filename, 'w') as f:
                json.dump(postdata, f, indent=4)
        return ""


    @cherrypy.expose
    def user_config_save(self, **postdata):
        if os.path.exists(postdata['zero_path']):
            save_user_config(postdata)

            #copy voice files to play via webpage
            print("copy voice files to play via webpage")
            sourcepath = os.path.join(postdata['zero_path'], 'data', 'se')
            destpath = os.path.join('.', 'voices')
            source_files = [ file for file in os.listdir(sourcepath) if file.startswith('ed7v') ]
            for file in source_files:
                src = os.path.join(sourcepath, file)
                dst = os.path.join(destpath, file)
                if not os.path.exists(dst):
                    print(file)
                    shutil.copyfile(src, dst)            
            raise cherrypy.HTTPRedirect('/public')
        else:
            raise cherrypy.HTTPRedirect('/public/config.html')


def webinterface(url='http://127.0.0.1:8080/lastvoices'):
    webbrowser.open(url)
    conf = {
        '/': {
            'tools.sessions.on': True,
            'tools.staticdir.root': os.path.abspath(os.getcwd())
        },
        '/public': {
            'tools.staticdir.on': True,
            'tools.staticdir.dir': 'public',
            'tools.staticdir.index': 'index.html'
        },
        '/media': {
            'tools.staticdir.on': True,
            'tools.staticdir.dir': 'media'
        },
        '/voices': {
            'tools.staticdir.on': True,
            'tools.staticdir.dir': 'voices'
        }          
    }
    cherrypy.quickstart(zeroVoice(),
                        '/',
                        conf)


def backup_translation_file():
    """backup translastion file"""
    now = datetime.datetime.now().strftime("_session_%Y%m%d_%H%M%S")
    shutil.copyfile(json_file, json_file + now)
    with open(json_file, 'r') as f:
        voices = json.load(f)


def load_user_config():
    try:
        if os.path.exists(user_config_file):
            with open(user_config_file, 'r') as f:
                user_config = json.load(f)
    except Exception as e:
        user_config = {}
    return user_config
    

def save_user_config(config):
    with open(user_config_file, 'w') as f:
        json.dump(config, f, indent = 4)


def init_voice_files():
    """load all voices"""
    files = [ file for file in os.listdir('./voices') if file.endswith('.wav') ]
    for file in files:
        voices[file] = voice_load(file)

        
def voice_load(voice):
    """load voice data from json file"""
    filename = './data/{voice}.json'.format(voice=voice[:-4])
    data = voice_data_dict.copy()
    data['voice'] = voice
    if os.path.exists(filename):
        with open(filename, 'r') as f:
            voice_data = json.load(f)
            data = {**data, **voice_data}
    return data


if __name__ == '__main__':

    user_config = load_user_config()
    pprint.pprint(user_config)
    #backup_translation_file()
    init_voice_files()

    #open event log
    seclog_available = None
    try:
        handle = win32evtlog.OpenEventLog(None, logtype)
        seclog_available = True
    except Exception as e:
        print('Eventlog could not be opened, administrator rights needed')
        seclog_available = False

    #split into two processes
    #process for the webinterface
    base_url = url='http://127.0.0.1:8080'
    zero_path = user_config.get('zero_path', '')
    if os.path.isdir(zero_path):
        if seclog_available == True:
            url = base_url + '/lastvoices'
        else:
            url = base_url + '/voices'
    else:
        url = base_url + '/public/config.html'
    webpid = threading.Thread(target=webinterface, kwargs=dict(url=url))
    webpid.start()
    
    #process collecting voice data in the background
    for voice in subscribe_and_yield_events(logtype, query=logqry):
        if voice:
            collect_voice(voice)

















