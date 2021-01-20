from .interfaces.icollisiondetectionmanager import ICollisionDetectionManager


class CollisionDetectionManager(ICollisionDetectionManager):
    def __init__(self, system_object):
        super().__init__(system_object)
