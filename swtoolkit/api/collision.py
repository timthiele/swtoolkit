from .interfaces.icollision import ICollision


class Collision(ICollision):
    def __init__(self, system_object):
        super().__init__(system_object)
