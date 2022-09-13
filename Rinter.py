from getResolve import get_resolve
from time import sleep
import utils

class Rinter:
    def __init__(self):
        self.project_manager = None
        self.project = None
        self.clip_index = None
        self.current_clip = None
        self.clips = None
        self.timeline = None
        self.resolve = None

    def start(self):
        successful = False
        while not successful:
            try:
                self.resolve = get_resolve()
                self.project_manager = self.resolve.GetProjectManager()
                self.project = self.project_manager.GetCurrentProject()
                self.timeline = self.project.GetCurrentTimeline()
                self.project.SetCurrentTimeline(self.timeline)
                self.clips = self.timeline.GetItemListInTrack("video", 1)
                self.current_clip = self.timeline.GetCurrentVideoItem()
                self.clip_index = utils.index_of_clip(
                    self.clips, self.current_clip)
                print("successfully loaded resolve")
                successful = True
            except ImportError:
                print("resolve api not found")
                exit(-1)
            except AttributeError:
                print("resolve not found. retrying in 5 seconds")
                sleep(5)
