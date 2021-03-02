from abc import ABCMeta, abstractmethod
import os


class ITask(metaclass=ABCMeta):
    """Interface of task."""
    def __init__(self, dir):
        self.dir = dir

    @abstractmethod
    def run(self):
        pass

    def is_valid_task(self):
        return os.path.exists(f"{os.getcwd()}/{self.dir}")
