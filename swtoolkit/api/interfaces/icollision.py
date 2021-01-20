import pythoncom
import win32com
from win32com.client import DispatchBaseClass
from win32com.client.dynamic import LCID


class ICollision(DispatchBaseClass):
    """Interface for Collision Detection"""
    # CLSID = IID('{900F5772-9B01-489B-AB74-C40586EA320A}')
    # coclass_clsid = IID('{987A2D3F-5135-4686-B5D4-4F842E0C9219}')

    def GetComponents(self):
        """Gets array of colliding components"""
        return self._ApplyTypes_(1, 1, (12, 0), (), 'GetComponents', None, )

    def GetTransforms(self):
        return self._ApplyTypes_(2, 1, (12, 0), (), 'GetTransforms', None, )

    def IsPenetrating(self):
        """Gets if collision is touch type or penetration type"""
        return self._oleobj_.InvokeTypes(3, LCID, 1, (11, 0), (), )

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
