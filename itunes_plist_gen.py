import datetime
import re
import json

import win32com.client
from sqlalchemy import *
from sqlalchemy.ext.declarative import declarative_base
# from sqlalchemy.orm import relation, sessionmaker
from sqlalchemy.orm import sessionmaker
import pickle
import os
import ConfigParser

usr_home = os.getenv('USERPROFILE')
if not usr_home:
    usr_home = os.path.expanduser('~/')
settings_dir = os.path.join(usr_home, '.itunes_plist_gen')
settings_file = os.path.join(settings_dir, 'settings.ini')
# make default directory and files if do not exist
if not os.path.exists(settings_dir):
    os.makedirs(settings_dir)
if not os.path.exists(settings_file):
    file = open(settings_file)
    file.write("""[Database]
db_engine: sqlite:///itunes.db
""")
    file.close()

config = ConfigParser.ConfigParser()
Config.read(settings_file)
db_engine = Config.get("Database",'db_engine')
print db_engine

engine = create_engine(db_engine)
Base = declarative_base(bind=engine)


class Track(Base):

    __tablename__ = 'tracks'
    PIDH = Column(Integer, primary_key=True)
    PIDL = Column(Integer, primary_key=True)
    Name = Column(String(500), nullable=True)
    Artist = Column(String(500), nullable=True)
    Album = Column(String(500), nullable=True)
    Kind = Column(String(250), nullable = True)
    DateAdded = Column(DateTime, nullable=False)
    PlayedDate = Column(DateTime, nullable=False)
    Grouping = Column(String(250), nullable = True)
    Year = Column(Integer, nullable = True)
    PlayedCount = Column(Integer)
    SkippedCount = Column(Integer)
    Score = Column(Float, nullable=False)
    IsMP3Copy = Column(Boolean, nullable=False)
    PIDHmp3 = Column(Integer, nullable = True)
    PIDLmp3 = Column(Integer, nullable = True)
    Enabled = Column(Boolean, nullable=False)
    Genre = Column(String(250), nullable = True)

    def __init__(self, itrack, iTunes):
        self.PIDH = iTunes.ITObjectPersistentIDHigh(itrack)
        self.PIDL = iTunes.ITObjectPersistentIDLow(itrack)
        self.update(itrack)
        self.IsMP3Copy = False

    def update(self, itrack):
        self.Name = itrack.Name
        self.Artist = itrack.Artist
        self.Album = itrack.Album

        self.DateAdded = datetime.datetime.fromtimestamp(int(itrack.DateAdded))
        self.Kind = itrack.KindAsString
        self.PlayedCount = itrack.PlayedCount

        self.Comment = itrack.Comment
        self.Grouping = itrack.Grouping
        self.Enabled = itrack.Enabled
        self.Year = itrack.Year
        self.Genre = itrack.Genre
        self.Score = 0
        if self.PlayedCount == 0:
            self.PlayedDate = datetime.datetime.fromtimestamp(0)
        else:
            try:
                self.PlayedDate = datetime.datetime.fromtimestamp(int(itrack.PlayedDate))
            except ValueError:
                self.PlayedDate = datetime.datetime.fromtimestamp(0)

        self.itrack = itrack
        m = re.search('mi:({.+?})', self.Comment)
        if m:
            d = json.loads(m.group(1))

            self.DateAdded = datetime.datetime.fromtimestamp(long(d['oau']));
        self.calculate_rating()
                #if "oau" not in track.Comment:
        #    track.Comment = "%s oau:[%d]" %(track.Comment,  actualtime);

    def get_itrack(self, iTunes):
        return iTunes.LibraryPlaylist.Tracks.ItemByPersistentID(self.PIDH, self.PIDL)

    def calculate_rating(self):
        timesince = datetime.datetime.now()-self.DateAdded
        #float(self.PlayedCount*100000000)/timesince.total_seconds()

        if timesince< datetime.timedelta(days=30)  and self.PlayedCount <3:
            self.Score=400
            return

        if self.PlayedCount ==0:
            self.Score=0
            return

        self.Score = float(self.PlayedCount*100000000)/timesince.total_seconds()

Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
s = Session()


def clear_playlist(playlist):
    tracks = playlist.Tracks
    while len(tracks) > 0:
        for track in tracks:
            try:
                print "deleting {0}".format(track.Name)
            except UnicodeEncodeError:
                pass

            track.Delete()

def label_purchased_tracks(iTunes):
    purchased_playlist_item = iTunes.LibrarySource.Playlists.ItemByName("Purchased")
    purchased_playlist = win32com.client.CastTo(purchased_playlist_item, 'IITLibraryPlaylist')
    for itrack in purchased_playlist:
        pass

def get_track(itrack, iTunes):
    track = s.query(Track).filter(Track.PIDH == iTunes.ITObjectPersistentIDHigh(itrack), Track.PIDL == iTunes.ITObjectPersistentIDLow(itrack)).first()
    if not track:
        track = Track(itrack, iTunes)
    else:
        track.update(itrack)
    s.add(track)

    return track

def load_all_tracks(iTunes):
    usrhome = os.getenv('USERPROFILE');
    if not usrhome:
       usrhome=os.path.expanduser('~/')
    settingsdir = os.path.join(usrhome, '.itunesman')
    fname = 'last_checked.pkl'
    cachefile= os.path.join(settingsdir, fname)
    try:
        pkl_file = open(cachefile, 'rb')
        last_checked = pickle.load(pkl_file)
        pkl_file.close()
    except IOError:
        last_checked = datetime.datetime.fromtimestamp(0)
        print "Warning: couldn't find cache file!"

    libraryTracks = iTunes.LibraryPlaylist.Tracks
    for itrack in libraryTracks:
        safe_playdate=1
        try:
            int(itrack.PlayedDate)
        except ValueError:
            safe_playdate =0
        if not safe_playdate or all(datetime.datetime.fromtimestamp(int(x)) > last_checked  for x in [ itrack.ModificationDate, safe_playdate  ]) :
            get_track(itrack, iTunes)


    s.commit()
    try:
        pkl_file = open(cachefile, 'wb')
        pickle.dump(datetime.datetime.now(), pkl_file)
        pkl_file.close()
    except IOError:
        last_checked = datetime.datetime.fromtimestamp(0)
def organize_mp3_copies(iTunes):
    need_mp3_conversion = get_playlist(iTunes, 'needmp3conversion')

    for track in s.query(Track).filter(Track.Kind != 'MPEG audio file').all():
        reptrack = s.query(Track).filter(Track.Name == track.Name, Track.Artist == track.Artist,  Track.Kind == 'MPEG audio file').first()

        if reptrack:
            ireptrack = reptrack.get_itrack(iTunes)
           # ireptrack.Enabled=False
            #ireptrack.PlayCount=0
            track.PIDHmp3 = reptrack.PIDH
            track.PIDLmp3 = reptrack.PIDH
            reptrack.IsMP3Copy =True
            reptrack.Score = 0
            reptrack.PlayCount = 0


        else:
            itrack = track.get_itrack(iTunes)
            #not working for some reason???/
            #if u'AAC' in track.Kind:
            #    needconversion.AddTrack(itrack)

    s.commit()
def get_playlist(iTunes, listname, do_clear_playlist=True):
    splaylist_item = iTunes.LibrarySource.Playlists.ItemByName(listname)
    if not splaylist_item:
        splaylist = win32com.client.CastTo(iTunes.CreatePlaylist(listname), 'IITLibraryPlaylist')
    else:
        splaylist = win32com.client.CastTo(splaylist_item, 'IITLibraryPlaylist')
        if do_clear_playlist:
            clear_playlist(splaylist)
    return splaylist

def create_list(iTunes, listname, query, replacewithMP3=False, make_backup=False):



    splaylist = get_playlist(iTunes, listname)
    if make_backup:
        splaylist_dated = get_playlist(iTunes, datetime.datetime.now().strftime(listname+"-%y%m%d"))


    for track in query:
        if 'Classical' in track.Genre or 'Soundtrack' in track.Genre or 'Intrumental' in track.Genre:
            continue
        if hasattr(track, 'Kind') and 'MPEG-4' in track.Kind:
            continue
        if replacewithMP3 and 'MPEG audio file' not in track.Kind:
            reptrack = s.query(Track).filter(Track.Name == track.Name, Track.Artist == track.Artist,  Track.Kind == 'MPEG audio file').first()
            if reptrack:
                track=reptrack
            else:
                try:
                    print 'no mp3 for {0}'.format(track.Name)
                except UnicodeEncodeError:
                    pass

                continue

        itrack = track.get_itrack(iTunes)
        #print track.Name

        try:
            if make_backup:
                splaylist_dated.AddTrack(itrack)
            splaylist.AddTrack(itrack)
        except:
            pass

        try:
            print track.Score, track.PlayedCount, track.Name, track.Genre
        except UnicodeEncodeError:
            pass


def main():
    iTunes = win32com.client.gencache.EnsureDispatch("iTunes.Application")
    load_all_tracks(iTunes)
    organize_mp3_copies(iTunes)
    create_list(iTunes, '$t1', s.query(Track).filter(Track.Enabled == True).order_by(Track.PlayedCount.desc()).limit(200))
    create_list(iTunes, '$s1', s.query(Track).filter(Track.Enabled == True).order_by(Track.Score.desc()).limit(200), make_backup=True)
    create_list(iTunes, '$sz00s', s.query(Track).filter(Track.Enabled == True).filter(Track.Year <=2010, Track.Year >=2000).order_by(Track.Score.desc()).limit(200))
    create_list(iTunes, '$sz90s', s.query(Track).filter(Track.Enabled == True).filter(Track.Year <=2000, Track.Year >=1990).order_by(Track.Score.desc()).limit(150))
    create_list(iTunes, '$sz80s', s.query(Track).filter(Track.Enabled == True).filter(Track.Year <=1990, Track.Year >=1980).order_by(Track.Score.desc()).limit(75))
    create_list(iTunes, '$t1m', s.query(Track).filter(Track.Enabled == True).order_by(Track.PlayedCount.desc()).limit(200), replacewithMP3=True)
    create_list(iTunes, '$s1m', s.query(Track).filter(Track.Enabled == True).order_by(Track.Score.desc()).limit(200), replacewithMP3=True)
    create_list(iTunes, '$u1', s.query(Track).filter(Track.Enabled == True).filter(Track.PlayedDate < (datetime.datetime.now() -  datetime.timedelta(days=60))).order_by(Track.Score.desc()).limit(200))
    create_list(iTunes, '$v1', s.query(Track).filter(Track.Enabled == True).filter(Track.PlayedDate < (datetime.datetime.now() -  datetime.timedelta(days=1))).order_by(Track.Score.desc()).limit(200))


if __name__ == "__main__":
    main()
