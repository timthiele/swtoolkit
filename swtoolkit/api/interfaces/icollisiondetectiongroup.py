import pythoncom
import win32com
from win32com.client import DispatchBaseClass
from win32com.client.dynamic import LCID

defaultNamedNotOptArg=pythoncom.Empty

class ICollisionDetectionGroup(DispatchBaseClass):
    """Interface for Collision Detection Group"""
    # CLSID = IID('{2DF28265-643B-48FD-920F-30F57131AABC}')
    # coclass_clsid = IID('{BD0006AC-D38F-4AD6-8822-5E2232841135}')

    def ApplyTransforms(self, ComponentTransforms=defaultNamedNotOptArg):
        """Apply transformations to components in group"""
        return self._oleobj_.InvokeTypes(4, LCID, 1, (3, 0), ((12, 1),),
                                         ComponentTransforms
                                         )

    def GetComponentCount(self):
        """Gets Number of Components in Group"""
        return self._oleobj_.InvokeTypes(2, LCID, 1, (3, 0), (), )

    def GetComponents(self):
        """Gets array of Components in group"""
        return self._ApplyTypes_(3, 1, (12, 0), (), 'GetComponents', None, )

    def GetTransforms(self):
        """Gets transforms for components in array"""
        return self._ApplyTypes_(5, 1, (12, 0), (), 'GetTransforms', None, )

    def RemoveAllTransforms(self):
        """Removes transforms from all the components in group"""
        return self._oleobj_.InvokeTypes(6, LCID, 1, (24, 0), (), )

    def SetComponents(self, Components=defaultNamedNotOptArg):
        """Sets the components for the Collision Detection group"""
        return self._oleobj_.InvokeTypes(1, LCID, 1, (3, 0), ((12, 1),),
                                         Components
                                         )

    _prop_map_get_ = {
    }
    _prop_map_put_ = {
    }

    def __iter__(self):
        """Return a Python iterator for this object"""
        try:
            ob = self._oleobj_.InvokeTypes(-4, LCID, 3, (13, 10), ())
        except pythoncom.error:
            raise TypeError("This object does not support enumeration")
        return win32com.client.util.Iterator(ob, None)
