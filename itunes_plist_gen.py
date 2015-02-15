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
db_location = os.path.join(settings_dir, 'itunes.db')
# make default directory and files if do not exist
if not os.path.exists(settings_dir):
    os.makedirs(settings_dir)
if not os.path.exists(settings_file):
    file1 = open(settings_file, 'w')
    file1.write("""[Database]
db_engine: sqlite:///{}
""".format(db_location))
    file1.close()

config = ConfigParser.ConfigParser()
config.read(settings_file)
db_engine = config.get("Database", 'db_engine')
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
    Kind = Column(String(250), nullable=True)
    DateAdded = Column(DateTime, nullable=False)
    PlayedDate = Column(DateTime, nullable=False)
    Comment = Column(String(500), nullable=False)
    Grouping = Column(String(250), nullable=True)
    Year = Column(Integer, nullable=True)
    PlayedCount = Column(Integer)
    SkippedCount = Column(Integer)
    Score = Column(Float, nullable=False)
    IsMP3Copy = Column(Boolean, nullable=False)
    PIDHmp3 = Column(Integer, nullable=True)
    PIDLmp3 = Column(Integer, nullable=True)
    Enabled = Column(Boolean, nullable=False)
    Genre = Column(String(250), nullable=True)

    def __init__(self, itunes_track, itunes_handle):
        self.PIDH = itunes_handle.ITObjectPersistentIDHigh(itunes_track)
        self.PIDL = itunes_handle.ITObjectPersistentIDLow(itunes_track)
        self.update(itunes_track)
        self.IsMP3Copy = False

    def update(self, itunes_track):
        self.Name = itunes_track.Name
        self.Artist = itunes_track.Artist
        self.Album = itunes_track.Album
        self.DateAdded = datetime.datetime.fromtimestamp(int(itunes_track.DateAdded))
        self.Kind = itunes_track.KindAsString
        self.PlayedCount = itunes_track.PlayedCount
        self.Comment = itunes_track.Comment
        self.Grouping = itunes_track.Grouping
        self.Enabled = itunes_track.Enabled
        self.Year = itunes_track.Year
        self.Genre = itunes_track.Genre
        self.Score = 0
        if self.PlayedCount == 0:
            self.PlayedDate = datetime.datetime.fromtimestamp(0)
        else:
            try:
                self.PlayedDate = datetime.datetime.fromtimestamp(int(itunes_track.PlayedDate))
            except ValueError:
                self.PlayedDate = datetime.datetime.fromtimestamp(0)

        self.itrack = itunes_track
        m = re.search('mi:({.+?})', self.Comment)
        if m:
            d = json.loads(m.group(1))
            self.DateAdded = datetime.datetime.fromtimestamp(long(d['oau']))
        self.calculate_rating()

    def get_itunes_track_handle(self, itunes_handle):
        return itunes_handle.LibraryPlaylist.Tracks.ItemByPersistentID(self.PIDH, self.PIDL)

    def calculate_rating(self):
        timesince = datetime.datetime.now()-self.DateAdded
        # Front load new songs
        if timesince < datetime.timedelta(days=30) and self.PlayedCount < 3:
            self.Score = 400
            return

        if self.PlayedCount == 0:
            self.Score = 0
            return

        self.Score = float(self.PlayedCount*100000000)/timesince.total_seconds()


def convert_itunes_date(itunes_date):
    try:
        return datetime.datetime.fromtimestamp(int(itunes_date))
    except ValueError:
        return datetime.datetime.fromtimestamp(0)


class ITunesObj:
    itunes_handle = None

    def __init__(self, itunes_handle):
        self.itunes_handle = itunes_handle

    @staticmethod
    def clear_playlist(playlist):
        tracks = playlist.Tracks
        while len(tracks) > 0:
            for track in tracks:
                try:
                    print "deleting {0}".format(track.Name)
                except UnicodeEncodeError:
                    pass
                track.Delete()

    def get_purchased_playlist(self):
        purchased_playlist_item = self.itunes_handle.LibrarySource.Playlists.ItemByName("Purchased")
        purchased_playlist = win32com.client.CastTo(purchased_playlist_item, 'IITLibraryPlaylist')
        return purchased_playlist

    def get_playlist(self,  playlist_name, do_clear_playlist=True):
        playlist_handle = self.itunes_handle.LibrarySource.Playlists.ItemByName(playlist_name)
        if not playlist_handle:
            playlist_handle = win32com.client.CastTo(self.itunes_handle.CreatePlaylist(playlist_name), 'IITLibraryPlaylist')
        else:
            playlist_handle = win32com.client.CastTo(playlist_handle, 'IITLibraryPlaylist')
        if do_clear_playlist:
            self.clear_playlist(playlist_handle)
        return playlist_handle

    def get_track_from_handle(self, itunes_track_handle):
        track = s.query(Track)\
            .filter(Track.PIDH == self.itunes_handle.ITObjectPersistentIDHigh(itunes_track_handle),
                    Track.PIDL == self.itunes_handle.ITObjectPersistentIDLow(itunes_track_handle)).first()
        return track

    def update_tracks(self):
        last_known_played_date = datetime.datetime.fromtimestamp(0)
        last_updated_track = s.query(Track).order_by(Track.PlayedDate.desc()).first()
        if last_updated_track:
            last_known_played_date = last_updated_track.PlayedDate

        library_tracks = self.itunes_handle.LibraryPlaylist.Tracks
        for itunes_track_handle in library_tracks:

            if convert_itunes_date(itunes_track_handle.ModificationDate) < last_known_played_date:
                continue
            if convert_itunes_date(itunes_track_handle.PlayedDate) < last_known_played_date:
                continue

            track = self.get_track_from_handle(itunes_track_handle)
            if not track:
                track = Track(itunes_track_handle, self.itunes_handle)
                s.add(track)
            else:
                track.update(itunes_track_handle)
        s.commit()

    def organize_mp3_copies(self):
        need_mp3_conversion_playlist = self.get_playlist('need_mp3_conversion')
        for track in s.query(Track).filter(Track.Kind != 'MPEG audio file').all():
            mp3track = s.query(Track)\
                .filter(Track.Name == track.Name, Track.Artist == track.Artist, Track.Kind == 'MPEG audio file').first()
            if mp3track:
                mp3track_handle= mp3track.get_itunes_track_handle(self.itunes_handle)
                track.PIDHmp3 = mp3track.PIDH
                track.PIDLmp3 = mp3track.PIDH
                mp3track.IsMP3Copy = True
                mp3track.Score = 0
                # mp3track_handle.PlayCount = 0
                mp3track_handle.Enabled = False
            else:
                track_handle = track.get_itunes_track_handle(self.itunes_handle)
                try:
                    if track_handle:
                        need_mp3_conversion_playlist.AddTrack(track_handle)
                except:
                    print('error on '+track.Name)

        s.commit()

    def create_list(self, list_name, query, replace_with_mp3=False, make_backup=False):
        playlist = self.get_playlist(list_name, do_clear_playlist=True)
        if make_backup:
            playlist_dated = self.get_playlist(datetime.datetime.now().strftime(list_name+"-%y%m%d"))
        for track in query:
            if 'Classical' in track.Genre or 'Soundtrack' in track.Genre or 'Instrumental' in track.Genre:
                continue
            if hasattr(track, 'Kind') and 'MPEG-4' in track.Kind:
                continue
            if replace_with_mp3 and 'MPEG audio file' not in track.Kind:
                replacement_track = s.query(Track)\
                    .filter(Track.Name == track.Name,
                            Track.Artist == track.Artist,
                            Track.Kind == 'MPEG audio file').first()
                if replacement_track:
                    track=replacement_track
                else:
                    try:
                        print 'no mp3 for {0}'.format(track.Name)
                    except UnicodeEncodeError:
                        pass
                    continue
            itunes_track_handle = track.get_itunes_track_handle(self.itunes_handle)

            if make_backup:
                playlist_dated.AddTrack(itunes_track_handle)
            try:
                playlist.AddTrack(itunes_track_handle)
                #pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2147024809), None)
            except Exception:
                print "error adding track " + track.Name + "by " + track.Artist


            try:
                print track.Score, track.PlayedCount, track.Name, track.Genre
            except UnicodeEncodeError:
                pass
Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
s = Session()

def main():
    i_tunes_instance = ITunesObj(win32com.client.gencache.EnsureDispatch("iTunes.Application"))
    i_tunes_instance.update_tracks()
    i_tunes_instance.organize_mp3_copies()
    i_tunes_instance.create_list('$t1', s.query(Track)
                                 .filter(Track.Enabled == True)
                                 .order_by(Track.PlayedCount.desc()).limit(200))
    i_tunes_instance.create_list('$s1', s.query(Track)
                                 .filter(Track.Enabled == True)
                                 .order_by(Track.Score.desc()).limit(200), make_backup=True)
    i_tunes_instance.create_list('$sz00s', s.query(Track)
                                 .filter(Track.Enabled == True)
                                 .filter(Track.Year <=2010, Track.Year >= 2000)
                                 .order_by(Track.Score.desc()).limit(200))
    i_tunes_instance.create_list('$sz90s', s.query(Track)
                                 .filter(Track.Enabled == True)
                                 .filter(Track.Year <=2000, Track.Year >= 1990)
                                 .order_by(Track.Score.desc()).limit(150))
    i_tunes_instance.create_list('$sz80s', s.query(Track).filter(Track.Enabled == True)
                                 .filter(Track.Year <=1990, Track.Year >=1980)
                                 .order_by(Track.Score.desc()).limit(75))
    i_tunes_instance.create_list('$t1m', s.query(Track).filter(Track.Enabled == True)
                                 .order_by(Track.PlayedCount.desc()).limit(200), replace_with_mp3=True)
    i_tunes_instance.create_list('$s1m', s.query(Track).filter(Track.Enabled == True)
                                 .order_by(Track.Score.desc()).limit(200), replace_with_mp3=True)
    i_tunes_instance.create_list('$u1', s.query(Track).filter(Track.Enabled == True)
                                 .filter(Track.PlayedDate < (datetime.datetime.now() - datetime.timedelta(days=60)))
                                 .order_by(Track.Score.desc()).limit(200))
    i_tunes_instance.create_list('$v1', s.query(Track)
                                 .filter(Track.Enabled == True)
                                 .filter(Track.PlayedDate < (datetime.datetime.now() - datetime.timedelta(days=1)))
                                 .order_by(Track.Score.desc()).limit(200))

if __name__ == "__main__":
    main()
