from abc import ABCMeta, abstractmethod


class ITask(metaclass=ABCMeta):
    """Interface of task."""
    def __init__(self, dir):
        self.dir = dir

    @abstractmethod
    def run(self):
        pass
