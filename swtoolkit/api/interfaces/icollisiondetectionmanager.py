import pythoncom
import win32com
from win32com.client import DispatchBaseClass
from win32com.client.dynamic import LCID, Dispatch

defaultNamedNotOptArg = pythoncom.Empty


class ICollisionDetectionManager(DispatchBaseClass):
    """Interface for Collision Detection Manager"""

    # CLSID = IID('{38175E7E-C4A3-43DC-94C1-CF14D29D8B2A}')
    # coclass_clsid = IID('{BE6CD5F5-C5A9-47BF-9682-D75044EBE2C0}')

    def CreateGroup(self):
        """Create Collision Detection Group Object"""
        ret = self._oleobj_.InvokeTypes(1, LCID, 1, (9, 0), (), )
        if ret is not None:
            ret = Dispatch(ret, 'CreateGroup', None)
        return ret

    def GetAssembly(self):
        """Gets Assembly Owner for CollisionDetection Manager """
        ret = self._oleobj_.InvokeTypes(8, LCID, 1, (9, 0), (), )
        if ret is not None:
            ret = Dispatch(ret, 'GetAssembly', None)
        return ret

    def GetCollisions(self, TreatContactAsCollision=defaultNamedNotOptArg,
                      Collisions=pythoncom.Missing):
        """Gets array of Collissions"""
        return self._ApplyTypes_(7, 1, (3, 0), ((11, 1), (16396, 2)),
                                 'GetCollisions', None, TreatContactAsCollision
                                 , Collisions)

    def GetGroup(self, GroupIndex=defaultNamedNotOptArg):
        """Gets group at specified index"""
        ret = self._oleobj_.InvokeTypes(3, LCID, 1, (9, 0), ((3, 1),),
                                        GroupIndex
                                        )
        if ret is not None:
            ret = Dispatch(ret, 'GetGroup', None)
        return ret

    def GetGroupCount(self):
        """Gets total number of groups"""
        return self._oleobj_.InvokeTypes(2, LCID, 1, (3, 0), (), )

    def IsCollisionPresent(self, TreatContactAsCollision=defaultNamedNotOptArg):
        """Checks if collision is present"""
        return self._oleobj_.InvokeTypes(6, LCID, 1, (3, 0), ((11, 1),),
                                         TreatContactAsCollision
                                         )

    def RemoveGroup(self, GroupIndex=defaultNamedNotOptArg):
        """Removes a group at specified index"""
        return self._oleobj_.InvokeTypes(4, LCID, 1, (11, 0), ((3, 1),),
                                         GroupIndex
                                         )

    def SetAssembly(self, OwnerAssem=defaultNamedNotOptArg):
        """Sets Assembly Owner for CollsionDetection Manager """
        return self._oleobj_.InvokeTypes(9, LCID, 1, (3, 0), ((9, 1),),
                                         OwnerAssem
                                         )

    _prop_map_get_ = {
        "GraphicsRedrawEnabled": (
            5, 2, (11, 0), (), "GraphicsRedrawEnabled", None),
    }
    _prop_map_put_ = {
        "GraphicsRedrawEnabled": ((5, LCID, 4, 0), ()),
    }

    def __iter__(self):
        """Return a Python iterator for this object"""
        try:
            ob = self._oleobj_.InvokeTypes(-4, LCID, 3, (13, 10), ())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)
