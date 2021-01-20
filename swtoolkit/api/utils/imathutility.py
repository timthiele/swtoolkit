import win32com.client
import pythoncom


class IMathUtility:
    def __init__(self, system_object):
        self.system_object = system_object

    @property
    def _instance(self):
        return self.system_object

    def compose_transform(self, xvec, yvec, zvec, transvec, scale):
        x_vector = win32com.client.VARIANT(pythoncom.VT_I4, xvec)
        y_vector = win32com.client.VARIANT(pythoncom.VT_I4, yvec)
        z_vector = win32com.client.VARIANT(pythoncom.VT_I4, zvec)
        trans_vector = win32com.client.VARIANT(pythoncom.VT_I4, transvec)
        scale_arg = win32com.client.VARIANT(pythoncom.VT_I4, scale)
        return self._instance.ComposeTransform(
            x_vector, y_vector, z_vector, trans_vector, scale_arg
        )
