from .interfaces.icollisiondetectiongroup import ICollisionDetectionGroup


class CollisionDetectionGroup(ICollisionDetectionGroup):
    def __init__(self, system_object):
        super().__init__(system_object)
