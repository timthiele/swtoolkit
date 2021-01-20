""" isldworks is a direct reimplementation of the ISldWorks interface
in the SolidWorks API.

http://help.solidworks.com/2020/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISldWorks.html
"""
import pythoncom
import win32com.client.CLSIDToClass
import win32com.client.util

from ..com import COM

# from pywintypes import IID
# from win32com.client import Dispatch, DispatchBaseClass

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg = pythoncom.Empty
defaultNamedNotOptArg = pythoncom.Empty
defaultUnnamedArg = pythoncom.Empty


# CLSID = IID('{83A33D31-27C5-11CE-BFD4-00400513BB57}')
# MajorVersion = 28
# MinorVersion = 0
# LibraryFlags = 8
# LCID = 0x0
#


class ISldWorks:
    def __init__(self):
        self._isldworks = COM("SldWorks.Application")

    @property
    def _instance(self):
        return self._isldworks

    def _active_doc(self):
        return self._instance.ActiveDoc

    def _get_visible(self):
        """Gets the visibility of the SolidWorks session."""
        return self._instance.Visible

    def _set_visible(self, state: bool):
        """Sets the visibility of the SolidWorks session.

        Args:
            state (bool): The visibility state. True is visible
        """
        self._instance.Visible = state

    def _get_frame_state(self):
        return self._instance.FrameState

    def _set_frame_state(self, state: int):
        self._instance.FrameState = state

    @property
    def startup_completed(self):
        return self._instance.StartupProcessCompleted

    def _opendoc6(
        self, filename: str, type_value: int, options: int, configuration: str
    ):
        """Opens a native solidworks document """

        arg1 = win32com.client.VARIANT(
            pythoncom.VT_BSTR, filename.replace("\\", "/")
        )
        arg2 = win32com.client.VARIANT(pythoncom.VT_I4, type_value)
        arg3 = win32com.client.VARIANT(pythoncom.VT_I4, options)
        arg4 = win32com.client.VARIANT(pythoncom.VT_BSTR, configuration)
        arg5 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None
        )
        arg6 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None
        )

        openDoc = self._instance.OpenDoc6
        openDoc(arg1, arg2, arg3, arg4, arg5, arg6)

        return arg5, arg6  # (Errors, Warnings)

    def activate_doc(self, *args, **kwargs):
        # Activates a loaded document and rebuilds it as specified.

        arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, args[0])
        arg2 = win32com.client.VARIANT(
            pythoncom.VT_BOOL, kwargs["use_user_preference"]
        )
        arg3 = win32com.client.VARIANT(pythoncom.VT_I4, kwargs["option"])
        arg4 = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, None
        )

        ActivateDoc = self._instance.ActivateDoc3
        ActivateDoc(arg1, arg2, arg3, arg4)

        return arg4

    def close_all_documents(self, include_unsaved: bool):
        """Closes all open documents

        :param include_unsaved: Include unsaved documents is function execution
        :type include_unsaved: bool
        :return: Execution feedback. True if successeful
        :rtype: bool
        """

        arg1 = win32com.client.VARIANT(pythoncom.VT_BOOL, include_unsaved)
        return self._instance.CloseAllDocuments(arg1)

    def close_doc(self, name):
        arg = win32com.client.VARIANT(pythoncom.VT_BSTR, name)
        return self._instance.CloseDoc(arg)

    def new_document(self, template_name, paper_size, width, height):
        pass

    def move_document(self, *args, **kwargs):
        pass

    def is_background_processing_complete(self, path):
        pass

    def load_file(self, file_name, arg_string, import_data, errors):
        pass

    def preview_doc(self):
        pass

    def quit_doc(self):
        pass

    def run_command(self):
        pass

    def run_macro(self):
        pass

    def send_msg_to_user(self):
        pass

    def save_settings(self):
        pass

    def get_cwd(self):
        return self._instance.GetCurrentWorkingDirectory()

    def _get_documents(self):
        return self._instance.GetDocuments

    def exit_app(self):
        self._instance.ExitApp()

    def activate_task_pane(self):
        pass

    def get_imodeler(self):
        return self._instance.GetModeler()

    def get_mass_properties(self):
        pass

    def get_user_unit(self):
        pass

    def get_template_sizes(self):
        pass

    def _get_process_id(self):
        return self._instance.GetProcessID

    def get_imathutility(self):
        return self._instance.IGetMathUtility()

    def get_collision_detection_manager(self):
        """Gets Collision Detection Manager"""
        ret = self._instance.GetCollisionDetectionManager()
        if ret is not None:
            ret = win32com.Dispatch(ret, 'GetCollisionDetectionManager', None)
        return ret

    def loadfile4(self):
        pass

# class ISldWorks(DispatchBaseClass):
#     """Interface for SOLIDWORKS"""
#
#     def __init__(self):
#         self._isldworks = COM("SldWorks.Application")
#
#     # CLSID = IID('{83A33D22-27C5-11CE-BFD4-00400513BB57}')
#     # coclass_clsid = IID('{D134B411-3689-497D-B2D7-A27CB1066648}')
#     @property
#     def _instance(self):
#         return self._isldworks
#
#     def _active_doc(self):
#         return self._instance.ActiveDoc
#
#     def _get_visible(self):
#         """Gets the visibility of the SolidWorks session."""
#         return self._instance.Visible
#
#     def _set_visible(self, state: bool):
#         """Sets the visibility of the SolidWorks session.
#
#         Args:
#             state (bool): The visibility state. True is visible
#         """
#         self._instance.Visible = state
#
#     def _get_frame_state(self):
#         return self._instance.FrameState
#
#     def _set_frame_state(self, state: int):
#         self._instance.FrameState = state
#
#     @property
#     def startup_completed(self):
#         return self._instance.StartupProcessCompleted
#
#     def ActivateDoc(self, Name=defaultNamedNotOptArg):
#         """Activates a document"""
#         ret = self._oleobj_.InvokeTypes(3, LCID, 1, (9, 0), ((8, 1),), Name
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'ActivateDoc', None)
#         return ret
#
#     def ActivateDoc2(self, Name=defaultNamedNotOptArg,
#                      Silent=defaultNamedNotOptArg,
#                      Errors=defaultNamedNotOptArg):
#         'Activates a document'
#         return self._ApplyTypes_(91, 1, (9, 0), ((8, 1), (11, 1), (16387, 3)),
#                                  'ActivateDoc2', None, Name
#                                  , Silent, Errors)
#
#     def ActivateDoc3(self, Name=defaultNamedNotOptArg,
#                      UseUserPreferences=defaultNamedNotOptArg,
#                      Option=defaultNamedNotOptArg,
#                      Errors=defaultNamedNotOptArg):
#         'Activates a document'
#         return self._ApplyTypes_(306, 1, (9, 0),
#                                  ((8, 1), (11, 1), (3, 1), (16387, 3)),
#                                  'ActivateDoc3', None, Name
#                                  , UseUserPreferences, Option, Errors)
#
#     def ActivateTaskPane(self, TaskPaneID=defaultNamedNotOptArg):
#         'Activate a specific task pane'
#         return self._oleobj_.InvokeTypes(236, LCID, 1, (11, 0), ((3, 1),),
#                                          TaskPaneID
#                                          )
#
#     def AddCallback(self, Cookie=defaultNamedNotOptArg,
#                     CallbackFunction=defaultNamedNotOptArg):
#         'Register a general perpose callback handler'
#         return self._oleobj_.InvokeTypes(222, LCID, 1, (24, 0),
#                                          ((3, 1), (8, 1)), Cookie
#                                          , CallbackFunction)
#
#     def AddFileOpenItem(self, CallbackFcnAndModule=defaultNamedNotOptArg,
#                         Description=defaultNamedNotOptArg):
#         'Adds an item in the Save As drop down, with a callback function'
#         return self._oleobj_.InvokeTypes(14, LCID, 1, (11, 0), ((8, 1), (8, 1)),
#                                          CallbackFcnAndModule
#                                          , Description)
#
#     def AddFileOpenItem2(self, Cookie=defaultNamedNotOptArg,
#                          MethodName=defaultNamedNotOptArg,
#                          Description=defaultNamedNotOptArg,
#                          Extension=defaultNamedNotOptArg):
#         'Adds an item in the Open file drop down, with a callback method'
#         return self._oleobj_.InvokeTypes(168, LCID, 1, (11, 0),
#                                          ((3, 1), (8, 1), (8, 1), (8, 1)),
#                                          Cookie
#                                          , MethodName, Description, Extension)
#
#     def AddFileOpenItem3(self, Cookie=defaultNamedNotOptArg,
#                          MethodName=defaultNamedNotOptArg,
#                          Description=defaultNamedNotOptArg,
#                          Extension=defaultNamedNotOptArg
#                          , OptionLabel=defaultNamedNotOptArg,
#                          OptionMethodName=defaultNamedNotOptArg):
#         'Adds an item in the Open file drop down, with a callback method'
#         return self._oleobj_.InvokeTypes(234, LCID, 1, (11, 0), (
#             (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1)), Cookie
#                                          , MethodName, Description, Extension,
#                                          OptionLabel, OptionMethodName
#                                          )
#
#     def AddFileSaveAsItem(self, CallbackFcnAndModule=defaultNamedNotOptArg,
#                           Description=defaultNamedNotOptArg,
#                           Type=defaultNamedNotOptArg):
#         'Adds an item in the Open drop down, with a callback function'
#         return self._oleobj_.InvokeTypes(15, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (3, 1)),
#                                          CallbackFcnAndModule
#                                          , Description, Type)
#
#     def AddFileSaveAsItem2(self, Cookie=defaultNamedNotOptArg,
#                            MethodName=defaultNamedNotOptArg,
#                            Description=defaultNamedNotOptArg,
#                            Extension=defaultNamedNotOptArg
#                            , DocumentType=defaultNamedNotOptArg):
#         'Adds an item in the Save as drop down, with a callback method'
#         return self._oleobj_.InvokeTypes(170, LCID, 1, (11, 0), (
#             (3, 1), (8, 1), (8, 1), (8, 1), (3, 1)), Cookie
#                                          , MethodName, Description, Extension,
#                                          DocumentType)
#
#     def AddItemToThirdPartyPopupMenu(self, RegisterId=defaultNamedNotOptArg,
#                                      DocType=defaultNamedNotOptArg,
#                                      Item=defaultNamedNotOptArg,
#                                      CallbackFcnAndModule=defaultNamedNotOptArg
#                                      , CustomName=defaultNamedNotOptArg,
#                                      HintString=defaultNamedNotOptArg,
#                                      BitmapFileName=defaultNamedNotOptArg,
#                                      MenuItemTypeOption=defaultNamedNotOptArg):
#         'Add item to third party popup menu'
#         return self._oleobj_.InvokeTypes(281, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (3, 1)),
#                                          RegisterId
#                                          , DocType, Item, CallbackFcnAndModule,
#                                          CustomName, HintString
#                                          , BitmapFileName, MenuItemTypeOption)
#
#     def AddItemToThirdPartyPopupMenu2(self, RegisterId=defaultNamedNotOptArg,
#                                       DocType=defaultNamedNotOptArg,
#                                       Item=defaultNamedNotOptArg,
#                                       Identifier=defaultNamedNotOptArg
#                                       , CallbackFunction=defaultNamedNotOptArg,
#                                       EnableFunction=defaultNamedNotOptArg,
#                                       CustomName=defaultNamedNotOptArg,
#                                       HintString=defaultNamedNotOptArg,
#                                       BitmapFileName=defaultNamedNotOptArg
#                                       ,
#                                       MenuItemTypeOption=defaultNamedNotOptArg):
#         'Add item to third party popup menu'
#         return self._oleobj_.InvokeTypes(300, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (8, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1),
#             (8, 1),
#             (3, 1)), RegisterId
#                                          , DocType, Item, Identifier,
#                                          CallbackFunction, EnableFunction
#                                          , CustomName, HintString,
#                                          BitmapFileName, MenuItemTypeOption)
#
#     def AddMenu(self, DocType=defaultNamedNotOptArg, Menu=defaultNamedNotOptArg,
#                 Position=defaultNamedNotOptArg):
#         'Add menus recursively'
#         return self._oleobj_.InvokeTypes(97, LCID, 1, (3, 0),
#                                          ((3, 1), (8, 1), (3, 1)), DocType
#                                          , Menu, Position)
#
#     def AddMenuItem(self, DocType=defaultNamedNotOptArg,
#                     Menu=defaultNamedNotOptArg, Postion=defaultNamedNotOptArg,
#                     CallbackModuleAndFcn=defaultNamedNotOptArg):
#         'Add a menu item'
#         return self._oleobj_.InvokeTypes(57, LCID, 1, (3, 0),
#                                          ((3, 1), (8, 1), (3, 1), (8, 1)),
#                                          DocType
#                                          , Menu, Postion, CallbackModuleAndFcn)
#
#     def AddMenuItem2(self, DocumentType=defaultNamedNotOptArg,
#                      Cookie=defaultNamedNotOptArg,
#                      MenuItem=defaultNamedNotOptArg,
#                      Position=defaultNamedNotOptArg
#                      , MenuCallback=defaultNamedNotOptArg,
#                      MenuEnableMethod=defaultNamedNotOptArg,
#                      HintString=defaultNamedNotOptArg):
#         'Add a menu item'
#         return self._oleobj_.InvokeTypes(147, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (8, 1), (3, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, MenuItem, Position,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString)
#
#     def AddMenuItem3(self, DocumentType=defaultNamedNotOptArg,
#                      Cookie=defaultNamedNotOptArg,
#                      MenuItem=defaultNamedNotOptArg,
#                      Position=defaultNamedNotOptArg
#                      , MenuCallback=defaultNamedNotOptArg,
#                      MenuEnableMethod=defaultNamedNotOptArg,
#                      HintString=defaultNamedNotOptArg,
#                      BitmapFilePath=defaultNamedNotOptArg):
#         'Add a menu item alongwith bitmap'
#         return self._oleobj_.InvokeTypes(213, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (8, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, MenuItem, Position,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, BitmapFilePath)
#
#     def AddMenuItem4(self, DocumentType=defaultNamedNotOptArg,
#                      Cookie=defaultNamedNotOptArg,
#                      MenuItem=defaultNamedNotOptArg,
#                      Position=defaultNamedNotOptArg
#                      , MenuCallback=defaultNamedNotOptArg,
#                      MenuEnableMethod=defaultNamedNotOptArg,
#                      HintString=defaultNamedNotOptArg,
#                      BitmapFilePath=defaultNamedNotOptArg):
#         'Add a menu item alongwith bitmap'
#         return self._oleobj_.InvokeTypes(255, LCID, 1, (3, 0), (
#             (3, 1), (3, 1), (8, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, MenuItem, Position,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, BitmapFilePath)
#
#     def AddMenuItem5(self, DocumentType=defaultNamedNotOptArg,
#                      Cookie=defaultNamedNotOptArg,
#                      MenuItem=defaultNamedNotOptArg,
#                      Position=defaultNamedNotOptArg
#                      , MenuCallback=defaultNamedNotOptArg,
#                      MenuEnableMethod=defaultNamedNotOptArg,
#                      HintString=defaultNamedNotOptArg,
#                      ImageList=defaultNamedNotOptArg):
#         'Adds a menu item and bitmap to the SolidWorks interface. '
#         return self._oleobj_.InvokeTypes(314, LCID, 1, (3, 0), (
#             (3, 1), (3, 1), (8, 1), (3, 1), (8, 1), (8, 1), (8, 1), (12, 1)),
#                                          DocumentType
#                                          , Cookie, MenuItem, Position,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, ImageList)
#
#     def AddMenuPopupItem(self, DocType=defaultNamedNotOptArg,
#                          SelType=defaultNamedNotOptArg,
#                          Item=defaultNamedNotOptArg,
#                          CallbackFcnAndModule=defaultNamedNotOptArg
#                          , CustomNames=defaultNamedNotOptArg):
#         'Add a menu item to a right mouse button menu'
#         return self._oleobj_.InvokeTypes(58, LCID, 1, (3, 0), (
#             (3, 1), (3, 1), (8, 1), (8, 1), (8, 1)), DocType
#                                          , SelType, Item, CallbackFcnAndModule,
#                                          CustomNames)
#
#     def AddMenuPopupItem2(self, DocumentType=defaultNamedNotOptArg,
#                           Cookie=defaultNamedNotOptArg,
#                           SelectType=defaultNamedNotOptArg,
#                           PopupItemName=defaultNamedNotOptArg
#                           , MenuCallback=defaultNamedNotOptArg,
#                           MenuEnableMethod=defaultNamedNotOptArg,
#                           HintString=defaultNamedNotOptArg,
#                           CustomNames=defaultNamedNotOptArg):
#         'Add a menu item to a right mouse button menu'
#         return self._oleobj_.InvokeTypes(172, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, SelectType, PopupItemName,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, CustomNames)
#
#     def AddMenuPopupItem3(self, DocumentType=defaultNamedNotOptArg,
#                           Cookie=defaultNamedNotOptArg,
#                           SelectType=defaultNamedNotOptArg,
#                           PopupItemName=defaultNamedNotOptArg
#                           , MenuCallback=defaultNamedNotOptArg,
#                           MenuEnableMethod=defaultNamedNotOptArg,
#                           HintString=defaultNamedNotOptArg,
#                           CustomNames=defaultNamedNotOptArg):
#         'Add a menu item to a right mouse button menu'
#         return self._oleobj_.InvokeTypes(256, LCID, 1, (3, 0), (
#             (3, 1), (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, SelectType, PopupItemName,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, CustomNames)
#
#     def AddMenuPopupItem4(self, DocumentType=defaultNamedNotOptArg,
#                           Cookie=defaultNamedNotOptArg,
#                           SelectType=defaultNamedNotOptArg,
#                           PopupItemName=defaultNamedNotOptArg
#                           , MenuCallback=defaultNamedNotOptArg,
#                           MenuEnableMethod=defaultNamedNotOptArg,
#                           HintString=defaultNamedNotOptArg,
#                           CustomNames=defaultNamedNotOptArg):
#         'Add a menu item to a right mouse button menu'
#         return self._oleobj_.InvokeTypes(302, LCID, 1, (3, 0), (
#             (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, SelectType, PopupItemName,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, CustomNames)
#
#     def AddToolbar(self, ModuleName=defaultNamedNotOptArg,
#                    Title=defaultNamedNotOptArg,
#                    SmallBitmapHandle=defaultNamedNotOptArg,
#                    LargeBitmapHandle=defaultNamedNotOptArg):
#         'Adds a toolbar'
#         return self._oleobj_.InvokeTypes(60, LCID, 1, (3, 0),
#                                          ((8, 1), (8, 1), (3, 1), (3, 1)),
#                                          ModuleName
#                                          , Title, SmallBitmapHandle,
#                                          LargeBitmapHandle)
#
#     def AddToolbar2(self, ModuleNameIn=defaultNamedNotOptArg,
#                     TitleIn=defaultNamedNotOptArg,
#                     SmallBitmapHandleIn=defaultNamedNotOptArg,
#                     LargeBitmapHandleIn=defaultNamedNotOptArg
#                     , MenuPosIn=defaultNamedNotOptArg,
#                     DecTemplateTypeIn=defaultNamedNotOptArg):
#         'Adds a toolbar'
#         return self._oleobj_.InvokeTypes(128, LCID, 1, (3, 0), (
#             (8, 1), (8, 1), (3, 1), (3, 1), (3, 1), (3, 1)), ModuleNameIn
#                                          , TitleIn, SmallBitmapHandleIn,
#                                          LargeBitmapHandleIn, MenuPosIn,
#                                          DecTemplateTypeIn
#                                          )
#
#     def AddToolbar3(self, Cookie=defaultNamedNotOptArg,
#                     Title=defaultNamedNotOptArg,
#                     SmallBitmapResourceID=defaultNamedNotOptArg,
#                     LargeBitmapResourceID=defaultNamedNotOptArg
#                     , MenuPositionForToolbar=defaultNamedNotOptArg,
#                     DocumentType=defaultNamedNotOptArg):
#         'Adds a toolbar'
#         return self._oleobj_.InvokeTypes(148, LCID, 1, (3, 0), (
#             (3, 1), (8, 1), (3, 1), (3, 1), (3, 1), (3, 1)), Cookie
#                                          , Title, SmallBitmapResourceID,
#                                          LargeBitmapResourceID,
#                                          MenuPositionForToolbar, DocumentType
#                                          )
#
#     def AddToolbar4(self, Cookie=defaultNamedNotOptArg,
#                     Title=defaultNamedNotOptArg,
#                     SmallBitmapImage=defaultNamedNotOptArg,
#                     LargeBitmapImage=defaultNamedNotOptArg
#                     , MenuPositionForToolbar=defaultNamedNotOptArg,
#                     DocumentType=defaultNamedNotOptArg):
#         'Adds a toolbar'
#         return self._oleobj_.InvokeTypes(188, LCID, 1, (3, 0), (
#             (3, 1), (8, 1), (8, 1), (8, 1), (3, 1), (3, 1)), Cookie
#                                          , Title, SmallBitmapImage,
#                                          LargeBitmapImage,
#                                          MenuPositionForToolbar, DocumentType
#                                          )
#
#     def AddToolbar5(self, Cookie=defaultNamedNotOptArg,
#                     Title=defaultNamedNotOptArg,
#                     ImageList=defaultNamedNotOptArg,
#                     MenuPositionForToolbar=defaultNamedNotOptArg
#                     , DocumentType=defaultNamedNotOptArg):
#         'Adds a toolbar '
#         return self._oleobj_.InvokeTypes(313, LCID, 1, (3, 0), (
#             (3, 1), (8, 1), (12, 1), (3, 1), (3, 1)), Cookie
#                                          , Title, ImageList,
#                                          MenuPositionForToolbar, DocumentType)
#
#     def AddToolbarCommand(self, ModuleName=defaultNamedNotOptArg,
#                           ToolbarId=defaultNamedNotOptArg,
#                           ToolbarIndex=defaultNamedNotOptArg,
#                           CommandString=defaultNamedNotOptArg):
#         'Adds a command to a toolbar'
#         return self._oleobj_.InvokeTypes(61, LCID, 1, (11, 0),
#                                          ((8, 1), (3, 1), (3, 1), (8, 1)),
#                                          ModuleName
#                                          , ToolbarId, ToolbarIndex,
#                                          CommandString)
#
#     def AddToolbarCommand2(self, Cookie=defaultNamedNotOptArg,
#                            ToolbarId=defaultNamedNotOptArg,
#                            ToolbarIndex=defaultNamedNotOptArg,
#                            ButtonCallback=defaultNamedNotOptArg
#                            , ButtonEnableMethod=defaultNamedNotOptArg,
#                            ToolTip=defaultNamedNotOptArg,
#                            HintString=defaultNamedNotOptArg):
#         'Adds a command to a toolbar'
#         return self._oleobj_.InvokeTypes(150, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1)), Cookie
#                                          , ToolbarId, ToolbarIndex,
#                                          ButtonCallback, ButtonEnableMethod,
#                                          ToolTip
#                                          , HintString)
#
#     def AllowFailedFeatureCreation(self, YesNo=defaultNamedNotOptArg):
#         'If an API creation of a feature fails, should the feature still be created?'
#         return self._oleobj_.InvokeTypes(84, LCID, 1, (11, 0), ((11, 1),), YesNo
#                                          )
#
#     def ArrangeIcons(self):
#         'Arrange Icons'
#         return self._oleobj_.InvokeTypes(31, LCID, 1, (24, 0), (), )
#
#     def ArrangeWindows(self, Style=defaultNamedNotOptArg):
#         'Arrange windows'
#         return self._oleobj_.InvokeTypes(32, LCID, 1, (24, 0), ((3, 1),), Style
#                                          )
#
#     def BlockSkinning(self):
#         'Block Skinning, avoid skinning a window'
#         return self._oleobj_.InvokeTypes(250, LCID, 1, (11, 0), (), )
#
#     def CallBack(self, CallBackFunc=defaultNamedNotOptArg,
#                  DefaultRetVal=defaultNamedNotOptArg,
#                  CallBackArgs=defaultNamedNotOptArg):
#         'Callback a function in an add-in DLL'
#         return self._oleobj_.InvokeTypes(75, LCID, 1, (3, 0),
#                                          ((8, 1), (3, 1), (8, 1)), CallBackFunc
#                                          , DefaultRetVal, CallBackArgs)
#
#     def CheckpointConvertedDocument(self, DocName=defaultNamedNotOptArg):
#         'Save the specified document, if it is read-only and has been converted by SOLIDWORKS'
#         return self._oleobj_.InvokeTypes(98, LCID, 1, (3, 0), ((8, 1),), DocName
#                                          )
#
#     def CloseAllDocuments(self, IncludeUnsaved=defaultNamedNotOptArg):
#         'Close All Open Documents'
#         return self._oleobj_.InvokeTypes(229, LCID, 1, (11, 0), ((11, 1),),
#                                          IncludeUnsaved
#                                          )
#
#     def CloseAndReopen(self, Doc=defaultNamedNotOptArg,
#                        Option=defaultNamedNotOptArg, NewDoc=pythoncom.Missing):
#         'Close and reopen document with references open in memory'
#         return self._ApplyTypes_(305, 1, (3, 0), ((9, 1), (3, 1), (16393, 2)),
#                                  'CloseAndReopen', None, Doc
#                                  , Option, NewDoc)
#
#     def CloseDoc(self, Name=defaultNamedNotOptArg):
#         'Closes the named document'
#         return self._oleobj_.InvokeTypes(7, LCID, 1, (24, 0), ((8, 1),), Name
#                                          )
#
#     def Command(self, Command=defaultNamedNotOptArg,
#                 Args=defaultNamedNotOptArg):
#         'Display the file new dialog'
#         return self._ApplyTypes_(190, 1, (12, 0), ((3, 1), (12, 1)), 'Command',
#                                  None, Command
#                                  , Args)
#
#     def CopyAppearance(self, Object=defaultNamedNotOptArg):
#         'Copy appearance from input object or selected object'
#         return self._oleobj_.InvokeTypes(307, LCID, 1, (11, 0), ((9, 1),),
#                                          Object
#                                          )
#
#     def CopyDocument(self, SourceDoc=defaultNamedNotOptArg,
#                      DestDoc=defaultNamedNotOptArg,
#                      FromChildren=defaultNamedNotOptArg,
#                      ToChildren=defaultNamedNotOptArg
#                      , Option=defaultNamedNotOptArg):
#         'Moves a document along with its specified dependents to the destination'
#         return self._oleobj_.InvokeTypes(185, LCID, 1, (3, 0), (
#             (8, 1), (8, 1), (12, 1), (12, 1), (3, 1)), SourceDoc
#                                          , DestDoc, FromChildren, ToChildren,
#                                          Option)
#
#     def CreateNewWindow(self):
#         'Create new window for the active window'
#         return self._oleobj_.InvokeTypes(30, LCID, 1, (24, 0), (), )
#
#     # Result is of type IPtnrPMPage
#     def CreatePMPage(self, DialogId=defaultNamedNotOptArg,
#                      Title=defaultNamedNotOptArg,
#                      Handler=defaultNamedNotOptArg):
#         'Create a page for display in the PropertyManager'
#         ret = self._oleobj_.InvokeTypes(203, LCID, 1, (9, 0),
#                                         ((3, 1), (8, 1), (9, 1)), DialogId
#                                         , Title, Handler)
#         if ret is not None:
#             ret = Dispatch(ret, 'CreatePMPage',
#                            '{2A586331-A56D-44C9-AA32-2868A96F044D}')
#         return ret
#
#     def CreatePropertyManagerPage(self, Title=defaultNamedNotOptArg,
#                                   Options=defaultNamedNotOptArg,
#                                   Handler=defaultNamedNotOptArg,
#                                   Errors=defaultNamedNotOptArg):
#         'Create a page for display in the PropertyManager'
#         return self._ApplyTypes_(163, 1, (9, 0),
#                                  ((8, 1), (3, 1), (9, 1), (16387, 3)),
#                                  'CreatePropertyManagerPage', None, Title
#                                  , Options, Handler, Errors)
#
#     def CreatePrunedModelArchive(self, PathName=defaultNamedNotOptArg,
#                                  ZipPathName=defaultNamedNotOptArg):
#         'Return an archive containing a pruned model'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(125, LCID, 1, (8, 0), ((8, 1), (8, 1)),
#                                          PathName
#                                          , ZipPathName)
#
#     # Result is of type ITaskpaneView
#     def CreateTaskpaneView(self, Bitmap=defaultNamedNotOptArg,
#                            ToolTip=defaultNamedNotOptArg,
#                            PHandler=defaultNamedNotOptArg):
#         'Add an applevel taskpane view.'
#         ret = self._oleobj_.InvokeTypes(197, LCID, 1, (9, 0),
#                                         ((16387, 1), (8, 1), (9, 1)), Bitmap
#                                         , ToolTip, PHandler)
#         if ret is not None:
#             ret = Dispatch(ret, 'CreateTaskpaneView',
#                            '{EDBBA0E9-B701-419E-A4AE-3409DBF12D40}')
#         return ret
#
#     # Result is of type ITaskpaneView
#     def CreateTaskpaneView2(self, Bitmap=defaultNamedNotOptArg,
#                             ToolTip=defaultNamedNotOptArg):
#         'Add an applevel taskpane view.'
#         ret = self._oleobj_.InvokeTypes(219, LCID, 1, (9, 0), ((8, 1), (8, 1)),
#                                         Bitmap
#                                         , ToolTip)
#         if ret is not None:
#             ret = Dispatch(ret, 'CreateTaskpaneView2',
#                            '{EDBBA0E9-B701-419E-A4AE-3409DBF12D40}')
#         return ret
#
#     # Result is of type ITaskpaneView
#     def CreateTaskpaneView3(self, ImageList=defaultNamedNotOptArg,
#                             ToolTip=defaultNamedNotOptArg):
#         'Add an applevel taskpane view'
#         ret = self._oleobj_.InvokeTypes(318, LCID, 1, (9, 0), ((12, 1), (8, 1)),
#                                         ImageList
#                                         , ToolTip)
#         if ret is not None:
#             ret = Dispatch(ret, 'CreateTaskpaneView3',
#                            '{EDBBA0E9-B701-419E-A4AE-3409DBF12D40}')
#         return ret
#
#     def DateCode(self):
#         'Returns the date code of the application'
#         return self._oleobj_.InvokeTypes(11, LCID, 1, (3, 0), (), )
#
#     def DefineAttribute(self, Name=defaultNamedNotOptArg):
#         'Makes an attribute definition'
#         ret = self._oleobj_.InvokeTypes(25, LCID, 1, (9, 0), ((8, 1),), Name
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'DefineAttribute', None)
#         return ret
#
#     def DisplayStatusBar(self, OnOff=defaultNamedNotOptArg):
#         'Display Status Bar'
#         return self._oleobj_.InvokeTypes(29, LCID, 1, (24, 0), ((11, 1),), OnOff
#                                          )
#
#     def DocumentVisible(self, Visible=defaultNamedNotOptArg,
#                         Type=defaultNamedNotOptArg):
#         'Control visibility of document on open'
#         return self._oleobj_.InvokeTypes(24, LCID, 1, (24, 0),
#                                          ((11, 1), (3, 1)), Visible
#                                          , Type)
#
#     def DragToolbarButton(self, SourceToolbar=defaultNamedNotOptArg,
#                           TargetToolbar=defaultNamedNotOptArg,
#                           SourceIndex=defaultNamedNotOptArg,
#                           TargetIndex=defaultNamedNotOptArg):
#         'Copies a button from one toolbar to another'
#         return self._oleobj_.InvokeTypes(221, LCID, 1, (24, 0),
#                                          ((3, 1), (3, 1), (3, 1), (3, 1)),
#                                          SourceToolbar
#                                          , TargetToolbar, SourceIndex,
#                                          TargetIndex)
#
#     def DragToolbarButtonFromCommandID(self, CommandID=defaultNamedNotOptArg,
#                                        TargetToolbar=defaultNamedNotOptArg,
#                                        TargetIndex=defaultNamedNotOptArg):
#         'Copies a button to a toolbar using a command id'
#         return self._oleobj_.InvokeTypes(254, LCID, 1, (3, 0),
#                                          ((3, 1), (3, 1), (3, 1)), CommandID
#                                          , TargetToolbar, TargetIndex)
#
#     def EnablePhotoWorksProgressiveRender(self, BEnable=defaultNamedNotOptArg):
#         'Enable or disable progressive rendering of PhotoWorks'
#         return self._oleobj_.InvokeTypes(268, LCID, 1, (24, 0), ((11, 1),),
#                                          BEnable
#                                          )
#
#     def EnableStereoDisplay(self, BEnable=defaultNamedNotOptArg):
#         'Enable stereoscopic display view'
#         return self._oleobj_.InvokeTypes(68, LCID, 1, (11, 0), ((11, 1),),
#                                          BEnable
#                                          )
#
#     # Result is of type IEnumDocuments
#     def EnumDocuments(self):
#         'Enumerates the documents'
#         ret = self._oleobj_.InvokeTypes(77, LCID, 1, (13, 0), (), )
#         if ret is not None:
#             # See if this IUnknown is really an IDispatch
#             try:
#                 ret = ret.QueryInterface(pythoncom.IID_IDispatch)
#             except pythoncom.error:
#                 return ret
#             ret = Dispatch(ret, 'EnumDocuments',
#                            '{83A33DB3-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IEnumDocuments2
#     def EnumDocuments2(self):
#         'Enumerates the documents'
#         ret = self._oleobj_.InvokeTypes(162, LCID, 1, (13, 0), (), )
#         if ret is not None:
#             # See if this IUnknown is really an IDispatch
#             try:
#                 ret = ret.QueryInterface(pythoncom.IID_IDispatch)
#             except pythoncom.error:
#                 return ret
#             ret = Dispatch(ret, 'EnumDocuments2',
#                            '{76D82D71-339A-4D1C-91A1-F6AC0CF9B625}')
#         return ret
#
#     def ExitApp(self):
#         'Terminates the application'
#         return self._oleobj_.InvokeTypes(6, LCID, 1, (24, 0), (), )
#
#     def ExportHoleWizardItem(self, StdToExport=defaultNamedNotOptArg,
#                              DestinationFolderPath=defaultNamedNotOptArg):
#         'Export HoleWizard Item'
#         return self._oleobj_.InvokeTypes(321, LCID, 1, (3, 0), ((8, 1), (8, 1)),
#                                          StdToExport
#                                          , DestinationFolderPath)
#
#     def ExportToolboxItem(self, StdToExport=defaultNamedNotOptArg,
#                           DestinationFolderPath=defaultNamedNotOptArg):
#         'Export Toolbox Item'
#         return self._oleobj_.InvokeTypes(323, LCID, 1, (3, 0), ((8, 1), (8, 1)),
#                                          StdToExport
#                                          , DestinationFolderPath)
#
#     def Frame(self):
#         'Gives a handle to the Application Frame'
#         ret = self._oleobj_.InvokeTypes(5, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'Frame', None)
#         return ret
#
#     def GetActiveConfigurationName(self, FilePathName=defaultNamedNotOptArg):
#         'Get the active configuration name'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(189, LCID, 1, (8, 0), ((8, 1),),
#                                          FilePathName
#                                          )
#
#     def GetAddInObject(self, Clsid=defaultNamedNotOptArg):
#         'Get a API object from an Add-In DLL'
#         ret = self._oleobj_.InvokeTypes(165, LCID, 1, (9, 0), ((8, 1),), Clsid
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetAddInObject', None)
#         return ret
#
#     def GetApplySelectionFilter(self):
#         'Get Apply Selection Filter Status'
#         return self._oleobj_.InvokeTypes(109, LCID, 1, (11, 0), (), )
#
#     def GetBuildNumbers(self, BaseVersion=pythoncom.Missing,
#                         CurrentVersion=pythoncom.Missing):
#         'Gets the base build version and specific build number of the application'
#         return self._ApplyTypes_(257, 1, (24, 0), ((16392, 2), (16392, 2)),
#                                  'GetBuildNumbers', None, BaseVersion
#                                  , CurrentVersion)
#
#     def GetBuildNumbers2(self, BaseVersion=pythoncom.Missing,
#                          CurrentVersion=pythoncom.Missing,
#                          HotFixes=pythoncom.Missing):
#         'Gets the base build version. specific build number.and hotfixes of the application'
#         return self._ApplyTypes_(303, 1, (24, 0),
#                                  ((16392, 2), (16392, 2), (16392, 2)),
#                                  'GetBuildNumbers2', None, BaseVersion
#                                  , CurrentVersion, HotFixes)
#
#     def GetButtonPosition(self, PointAt=defaultNamedNotOptArg,
#                           LocX=pythoncom.Missing, LocY=pythoncom.Missing):
#         'Get screen coordinate of command button'
#         return self._ApplyTypes_(284, 1, (11, 0),
#                                  ((3, 1), (16387, 2), (16387, 2)),
#                                  'GetButtonPosition', None, PointAt
#                                  , LocX, LocY)
#
#     def GetCollisionDetectionManager(self):
#         'Gets Collision Detection Manager'
#         ret = self._oleobj_.InvokeTypes(329, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetCollisionDetectionManager', None)
#         return ret
#
#     def GetColorTable(self):
#         'Get color table'
#         ret = self._oleobj_.InvokeTypes(139, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetColorTable', None)
#         return ret
#
#     def GetCommandID(self, Clsid=defaultNamedNotOptArg,
#                      UserCmdID=defaultNamedNotOptArg):
#         "Get the SOLIDWORKS commandID for an addin's callback"
#         return self._oleobj_.InvokeTypes(230, LCID, 1, (3, 0), ((8, 1), (3, 1)),
#                                          Clsid
#                                          , UserCmdID)
#
#     # Result is of type ICommandManager
#     def GetCommandManager(self, Cookie=defaultNamedNotOptArg):
#         'Gets the command manager for a given addin'
#         ret = self._oleobj_.InvokeTypes(220, LCID, 1, (9, 0), ((3, 1),), Cookie
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetCommandManager',
#                            '{F61069CF-2E42-4AC4-A517-6A95B79E45EE}')
#         return ret
#
#     def GetConfigurationCount(self, FilePathName=defaultNamedNotOptArg):
#         'Get Configuration Count'
#         return self._oleobj_.InvokeTypes(178, LCID, 1, (3, 0), ((8, 1),),
#                                          FilePathName
#                                          )
#
#     def GetConfigurationNames(self, FilePathName=defaultNamedNotOptArg):
#         'Get Configuration Names'
#         return self._ApplyTypes_(179, 1, (12, 0), ((8, 1),),
#                                  'GetConfigurationNames', None, FilePathName
#                                  )
#
#     def GetCookie(self, AddinClsid=defaultNamedNotOptArg,
#                   ResourceModuleHandle=defaultNamedNotOptArg,
#                   AddinCallbacks=defaultNamedNotOptArg):
#         'Add old MFC type of addin to addin manager, useful for using command mgr APIs which require cookie'
#         return self._oleobj_.InvokeTypes(235, LCID, 1, (3, 0),
#                                          ((8, 1), (3, 1), (9, 1)), AddinClsid
#                                          , ResourceModuleHandle, AddinCallbacks)
#
#     def GetCookiex64(self, AddinClsid=defaultNamedNotOptArg,
#                      ResourceModuleHandle=defaultNamedNotOptArg,
#                      AddinCallbacks=defaultNamedNotOptArg):
#         'Add old MFC type of addin to addin manager, useful for using command mgr APIs which require cookie'
#         return self._oleobj_.InvokeTypes(292, LCID, 1, (3, 0),
#                                          ((8, 1), (20, 1), (9, 1)), AddinClsid
#                                          , ResourceModuleHandle, AddinCallbacks)
#
#     def GetCurrSolidWorksRegSubKey(self):
#         'Returns the name of current SOLIDWORKS Registry scope'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(117, LCID, 1, (8, 0), (), )
#
#     def GetCurrentFileUser(self, FilePathName=defaultNamedNotOptArg):
#         'Gets the current user name of a specified file'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(266, LCID, 1, (8, 0), ((8, 1),),
#                                          FilePathName
#                                          )
#
#     def GetCurrentKernelVersions(self, Version1=pythoncom.Missing,
#                                  Version2=pythoncom.Missing,
#                                  Version3=pythoncom.Missing):
#         'Gets the versions of current kernel engines'
#         return self._ApplyTypes_(124, 1, (24, 0),
#                                  ((16392, 2), (16392, 2), (16392, 2)),
#                                  'GetCurrentKernelVersions', None, Version1
#                                  , Version2, Version3)
#
#     def GetCurrentLanguage(self):
#         'Gets the current language'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(94, LCID, 1, (8, 0), (), )
#
#     def GetCurrentLicenseType(self):
#         return self._oleobj_.InvokeTypes(316, LCID, 1, (3, 0), (), )
#
#     def GetCurrentMacroPathFolder(self):
#         'Gets the current running Macros folder'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(226, LCID, 1, (8, 0), (), )
#
#     def GetCurrentMacroPathName(self):
#         'Gets the current running Macro pathname'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(121, LCID, 1, (8, 0), (), )
#
#     def GetCurrentWorkingDirectory(self):
#         'Gets the current working directory'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(86, LCID, 1, (8, 0), (), )
#
#     def GetDataFolder(self, BShowErrorMsg=defaultNamedNotOptArg):
#         'Returns the data folder name'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(88, LCID, 1, (8, 0), ((11, 1),),
#                                          BShowErrorMsg
#                                          )
#
#     def GetDocumentCount(self):
#         'Get the number of model Documents open in SOLIDWORKS'
#         return self._oleobj_.InvokeTypes(272, LCID, 1, (3, 0), (), )
#
#     def GetDocumentDependencies(self, Document=defaultNamedNotOptArg,
#                                 Traverseflag=defaultNamedNotOptArg,
#                                 Searchflag=defaultNamedNotOptArg):
#         'Return names of documents that this document references.'
#         return self._ApplyTypes_(70, 1, (12, 0), ((8, 1), (3, 1), (3, 1)),
#                                  'GetDocumentDependencies', None, Document
#                                  , Traverseflag, Searchflag)
#
#     def GetDocumentDependencies2(self, Document=defaultNamedNotOptArg,
#                                  Traverseflag=defaultNamedNotOptArg,
#                                  Searchflag=defaultNamedNotOptArg,
#                                  AddReadOnlyInfo=defaultNamedNotOptArg):
#         'Return names of documents that this document references.'
#         return self._ApplyTypes_(104, 1, (12, 0),
#                                  ((8, 1), (11, 1), (11, 1), (11, 1)),
#                                  'GetDocumentDependencies2', None, Document
#                                  , Traverseflag, Searchflag, AddReadOnlyInfo)
#
#     def GetDocumentDependenciesCount(self, Document=defaultNamedNotOptArg,
#                                      Traverseflag=defaultNamedNotOptArg,
#                                      Searchflag=defaultNamedNotOptArg):
#         'Returns the size of array needed for a call to IGetDocumentDependencies'
#         return self._oleobj_.InvokeTypes(72, LCID, 1, (3, 0),
#                                          ((8, 1), (3, 1), (3, 1)), Document
#                                          , Traverseflag, Searchflag)
#
#     def GetDocumentTemplate(self, Mode=defaultNamedNotOptArg,
#                             TemplateName=defaultNamedNotOptArg,
#                             PaperSize=defaultNamedNotOptArg,
#                             Width=defaultNamedNotOptArg
#                             , Height=defaultNamedNotOptArg):
#         'Displays new doc dialog and returns the name of selected document template'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(113, LCID, 1, (8, 0), (
#             (3, 1), (8, 1), (3, 1), (5, 1), (5, 1)), Mode
#                                          , TemplateName, PaperSize, Width,
#                                          Height)
#
#     def GetDocumentVisible(self, Type=defaultNamedNotOptArg):
#         'Get the visibility of document to open'
#         return self._oleobj_.InvokeTypes(244, LCID, 1, (11, 0), ((3, 1),), Type
#                                          )
#
#     def GetDocuments(self):
#         'Get the open model Documents in SOLIDWORKS'
#         return self._ApplyTypes_(273, 1, (12, 0), (), 'GetDocuments', None, )
#
#     def GetEdition(self):
#         'Get SOLIDWORKS edition'
#         return self._oleobj_.InvokeTypes(183, LCID, 1, (3, 0), (), )
#
#     def GetEnvironment(self):
#         'Get the SW environment'
#         ret = self._oleobj_.InvokeTypes(36, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetEnvironment', None)
#         return ret
#
#     def GetErrorMessages(self, Msgs=pythoncom.Missing, MsgIDs=pythoncom.Missing,
#                          MsgTypes=pythoncom.Missing):
#         'Get errror messages'
#         return self._ApplyTypes_(225, 1, (3, 0),
#                                  ((16396, 2), (16396, 2), (16396, 2)),
#                                  'GetErrorMessages', None, Msgs
#                                  , MsgIDs, MsgTypes)
#
#     def GetExecutablePath(self):
#         'Get the path for sldworks.exe for this app'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(182, LCID, 1, (8, 0), (), )
#
#     def GetExportFileData(self, FileType=defaultNamedNotOptArg):
#         'Gets the data for export to the given file type'
#         ret = self._oleobj_.InvokeTypes(237, LCID, 1, (9, 0), ((3, 1),),
#                                         FileType
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetExportFileData', None)
#         return ret
#
#     def GetFirstDocument(self):
#         'Get the first document in the session'
#         ret = self._oleobj_.InvokeTypes(85, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetFirstDocument', None)
#         return ret
#
#     def GetHoleStandardsData(self, HoleTypeID=defaultNamedNotOptArg):
#         'Get Hole Standards Data'
#         ret = self._oleobj_.InvokeTypes(328, LCID, 1, (9, 0), ((3, 1),),
#                                         HoleTypeID
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetHoleStandardsData', None)
#         return ret
#
#     def GetImageSize(self, Small=pythoncom.Missing, Medium=pythoncom.Missing,
#                      Large=pythoncom.Missing):
#         'Get Image Size required for current DPI setting '
#         return self._ApplyTypes_(317, 1, (3, 0),
#                                  ((16387, 2), (16387, 2), (16387, 2)),
#                                  'GetImageSize', None, Small
#                                  , Medium, Large)
#
#     def GetImportFileData(self, FileName=defaultNamedNotOptArg):
#         'Gets the IGES import data for the given file'
#         ret = self._oleobj_.InvokeTypes(217, LCID, 1, (9, 0), ((8, 1),),
#                                         FileName
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetImportFileData', None)
#         return ret
#
#     def GetInterfaceBrightnessThemeColors(self, Colors=pythoncom.Missing):
#         'Get Colors Array and current Interface Brightness Theme. '
#         return self._ApplyTypes_(315, 1, (3, 0), ((16396, 2),),
#                                  'GetInterfaceBrightnessThemeColors', None,
#                                  Colors
#                                  )
#
#     def GetLastSaveError(self, FilePath=pythoncom.Missing,
#                          ErrorCode=pythoncom.Missing):
#         'Get the last save error'
#         return self._ApplyTypes_(279, 1, (12, 0), ((16396, 2), (16396, 2)),
#                                  'GetLastSaveError', None, FilePath
#                                  , ErrorCode)
#
#     def GetLastToolbarID(self):
#         'Get the command ID of the last added toolbar command'
#         return self._oleobj_.InvokeTypes(202, LCID, 1, (3, 0), (), )
#
#     def GetLatestSupportedFileVersion(self):
#         'Gets the latest supported file version'
#         return self._oleobj_.InvokeTypes(214, LCID, 1, (3, 0), (), )
#
#     def GetLineStyles(self, StyleFile=defaultNamedNotOptArg,
#                       StyleNameList=pythoncom.Missing,
#                       StyleList=pythoncom.Missing):
#         'Get the line styles from a file'
#         return self._ApplyTypes_(295, 1, (11, 0),
#                                  ((8, 1), (16396, 2), (16396, 2)),
#                                  'GetLineStyles', None, StyleFile
#                                  , StyleNameList, StyleList)
#
#     def GetLocalizedMenuName(self, MenuId=defaultNamedNotOptArg):
#         'Returns a localized menu name for the given menu Id'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(103, LCID, 1, (8, 0), ((3, 1),), MenuId
#                                          )
#
#     def GetMacroMethods(self, FilePathName=defaultNamedNotOptArg,
#                         Filter=defaultNamedNotOptArg):
#         'Get a list of methods from a macro file'
#         return self._ApplyTypes_(267, 1, (12, 0), ((8, 1), (3, 1)),
#                                  'GetMacroMethods', None, FilePathName
#                                  , Filter)
#
#     def GetMassProperties(self, FilePathName=defaultNamedNotOptArg,
#                           ConfigurationName=defaultNamedNotOptArg):
#         'Gets the mass properties from the given document for a given configuration'
#         return self._ApplyTypes_(101, 1, (12, 0), ((8, 1), (8, 1)),
#                                  'GetMassProperties', None, FilePathName
#                                  , ConfigurationName)
#
#     def GetMassProperties2(self, FilePathName=defaultNamedNotOptArg,
#                            ConfigurationName=defaultNamedNotOptArg,
#                            Accuracy=defaultNamedNotOptArg):
#         'Gets the mass properties from the given document for a given configuration'
#         return self._ApplyTypes_(174, 1, (12, 0), ((8, 1), (8, 1), (3, 1)),
#                                  'GetMassProperties2', None, FilePathName
#                                  , ConfigurationName, Accuracy)
#
#     def GetMaterialDatabaseCount(self):
#         'Get the count of the material databases.'
#         return self._oleobj_.InvokeTypes(195, LCID, 1, (3, 0), (), )
#
#     def GetMaterialDatabases(self):
#         'Get the list of the material databases.'
#         return self._ApplyTypes_(194, 1, (12, 0), (), 'GetMaterialDatabases',
#                                  None, )
#
#     def GetMaterialSchemaPathName(self):
#         'Get pathname of XML schema of material properties.'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(193, LCID, 1, (8, 0), (), )
#
#     def GetMathUtility(self):
#         'Gets Math Utility Interface'
#         ret = self._oleobj_.InvokeTypes(132, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetMathUtility', None)
#         return ret
#
#     def GetMenuStrings(self, CommandID=defaultNamedNotOptArg,
#                        DocumentType=defaultNamedNotOptArg,
#                        ParentMenuName=pythoncom.Missing):
#         'Retrieves the label of the specified menu item'
#         return self._ApplyTypes_(240, 1, (8, 0), ((3, 1), (3, 1), (16392, 2)),
#                                  'GetMenuStrings', None, CommandID
#                                  , DocumentType, ParentMenuName)
#
#     # Result is of type IModelView
#     def GetModelView(self, ModelName=defaultNamedNotOptArg,
#                      WindowID=defaultNamedNotOptArg, Row=defaultNamedNotOptArg,
#                      Column=defaultNamedNotOptArg):
#         'Get the specified model view'
#         ret = self._oleobj_.InvokeTypes(275, LCID, 1, (9, 0),
#                                         ((8, 1), (3, 1), (3, 1), (3, 1)),
#                                         ModelName
#                                         , WindowID, Row, Column)
#         if ret is not None:
#             ret = Dispatch(ret, 'GetModelView',
#                            '{83A33D4C-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     def GetModeler(self):
#         'Get the Geometry Modeler'
#         ret = self._oleobj_.InvokeTypes(34, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetModeler', None)
#         return ret
#
#     def GetMouseDragMode(self, Command=defaultNamedNotOptArg):
#         'Is the specified command the currently running command?'
#         return self._oleobj_.InvokeTypes(93, LCID, 1, (11, 0), ((3, 1),),
#                                          Command
#                                          )
#
#     def GetOpenDocSpec(self, FileName=defaultNamedNotOptArg):
#         'Get interface for specifying a document'
#         ret = self._oleobj_.InvokeTypes(248, LCID, 1, (9, 0), ((8, 1),),
#                                         FileName
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetOpenDocSpec', None)
#         return ret
#
#     # Result is of type IModelDoc2
#     def GetOpenDocument(self, DocName=defaultNamedNotOptArg):
#         'Gets the open document with the given title/path name'
#         ret = self._oleobj_.InvokeTypes(216, LCID, 1, (9, 0), ((8, 1),), DocName
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetOpenDocument',
#                            '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}')
#         return ret
#
#     def GetOpenDocumentByName(self, DocumentName=defaultNamedNotOptArg):
#         'Gets the open document with the given name'
#         ret = self._oleobj_.InvokeTypes(122, LCID, 1, (9, 0), ((8, 1),),
#                                         DocumentName
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetOpenDocumentByName', None)
#         return ret
#
#     def GetOpenFileName(self, DialogTitle=defaultNamedNotOptArg,
#                         InitialFileName=defaultNamedNotOptArg,
#                         FileFilter=defaultNamedNotOptArg,
#                         OpenOptions=pythoncom.Missing
#                         , ConfigName=pythoncom.Missing,
#                         DisplayName=pythoncom.Missing):
#         'Ask user for a file name'
#         return self._ApplyTypes_(211, 1, (8, 0), (
#             (8, 1), (8, 1), (8, 1), (16387, 2), (16392, 2), (16392, 2)),
#                                  'GetOpenFileName', None, DialogTitle
#                                  , InitialFileName, FileFilter, OpenOptions,
#                                  ConfigName, DisplayName
#                                  )
#
#     def GetOpenedFileInfo(self, FileName=pythoncom.Missing,
#                           Options=pythoncom.Missing):
#         'Gets file open and its open options'
#         return self._ApplyTypes_(215, 1, (24, 0), ((16392, 2), (16387, 2)),
#                                  'GetOpenedFileInfo', None, FileName
#                                  , Options)
#
#     def GetPreviewBitmap(self, FilePathName=defaultNamedNotOptArg,
#                          ConfigName=defaultNamedNotOptArg):
#         'Get Configuration Picture'
#         ret = self._oleobj_.InvokeTypes(181, LCID, 1, (9, 0), ((8, 1), (8, 1)),
#                                         FilePathName
#                                         , ConfigName)
#         if ret is not None:
#             ret = Dispatch(ret, 'GetPreviewBitmap', None)
#         return ret
#
#     def GetPreviewBitmapFile(self, DocumentPath=defaultNamedNotOptArg,
#                              ConfigName=defaultNamedNotOptArg,
#                              BitMapFile=defaultNamedNotOptArg):
#         'Get a preview bitmap and save to a filename'
#         return self._oleobj_.InvokeTypes(253, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (8, 1)), DocumentPath
#                                          , ConfigName, BitMapFile)
#
#     def GetProcessID(self):
#         'Get the process ID for this process'
#         return self._oleobj_.InvokeTypes(166, LCID, 1, (3, 0), (), )
#
#     def GetRayTraceRenderer(self, RendererType=defaultNamedNotOptArg):
#         'Get the ray trace renderer'
#         ret = self._oleobj_.InvokeTypes(296, LCID, 1, (9, 0), ((3, 1),),
#                                         RendererType
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetRayTraceRenderer', None)
#         return ret
#
#     def GetRecentFiles(self):
#         'Get a comma separated list of the Most Recent Used files.'
#         return self._ApplyTypes_(191, 1, (12, 0), (), 'GetRecentFiles', None, )
#
#     # Result is of type IRoutingSettings
#     def GetRoutingSettings(self):
#         'Get routing settings'
#         ret = self._oleobj_.InvokeTypes(294, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetRoutingSettings',
#                            '{909734EF-D5AD-4FFD-84AC-05A320A9F349}')
#         return ret
#
#     def GetRunningCommandInfo(self, CommandID=pythoncom.Missing,
#                               PMTitle=pythoncom.Missing,
#                               IsUiActive=pythoncom.Missing):
#         'Get Running Command information'
#         return self._ApplyTypes_(293, 1, (24, 0),
#                                  ((16387, 2), (16392, 2), (16395, 2)),
#                                  'GetRunningCommandInfo', None, CommandID
#                                  , PMTitle, IsUiActive)
#
#     def GetSSOFormattedURL(self, TargetUrl=defaultNamedNotOptArg):
#         'Gets the Single Signon formatted version of the URL'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(327, LCID, 1, (8, 0), ((8, 1),),
#                                          TargetUrl
#                                          )
#
#     def GetSafeArrayUtility(self):
#         'Get an instance of the safe array utilit object'
#         ret = self._oleobj_.InvokeTypes(309, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetSafeArrayUtility', None)
#         return ret
#
#     def GetSaveTo3DExperienceOptions(self):
#         'Get 3DExperience save options'
#         ret = self._oleobj_.InvokeTypes(331, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetSaveTo3DExperienceOptions', None)
#         return ret
#
#     def GetSearchFolders(self, FolderType=defaultNamedNotOptArg):
#         'Get Search Folders'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(42, LCID, 1, (8, 0), ((3, 1),),
#                                          FolderType
#                                          )
#
#     def GetSelectionFilter(self, SelType=defaultNamedNotOptArg):
#         'Get selection filter status'
#         return self._oleobj_.InvokeTypes(89, LCID, 1, (11, 0), ((3, 1),),
#                                          SelType
#                                          )
#
#     def GetSelectionFilters(self):
#         'Get all selection filters with status=ON'
#         return self._ApplyTypes_(107, 1, (12, 0), (), 'GetSelectionFilters',
#                                  None, )
#
#     def GetTemplateSizes(self, FileName=defaultNamedNotOptArg):
#         'Get the sheet properties from a template document'
#         return self._ApplyTypes_(137, 1, (12, 0), ((8, 1),), 'GetTemplateSizes',
#                                  None, FileName
#                                  )
#
#     def GetToolbarDock(self, ModuleIn=defaultNamedNotOptArg,
#                        ToolbarIDIn=defaultNamedNotOptArg):
#         'Get the docking state of the toolbar'
#         return self._oleobj_.InvokeTypes(130, LCID, 1, (3, 0), ((8, 1), (3, 1)),
#                                          ModuleIn
#                                          , ToolbarIDIn)
#
#     def GetToolbarDock2(self, Cookie=defaultNamedNotOptArg,
#                         ToolbarId=defaultNamedNotOptArg):
#         'Get the docking state of the toolbar'
#         return self._oleobj_.InvokeTypes(154, LCID, 1, (3, 0), ((3, 1), (3, 1)),
#                                          Cookie
#                                          , ToolbarId)
#
#     def GetToolbarState(self, Module=defaultNamedNotOptArg,
#                         ToolbarId=defaultNamedNotOptArg,
#                         ToolbarState=defaultNamedNotOptArg):
#         'Gets the state of the toolbar'
#         return self._oleobj_.InvokeTypes(65, LCID, 1, (11, 0),
#                                          ((8, 1), (3, 1), (3, 1)), Module
#                                          , ToolbarId, ToolbarState)
#
#     def GetToolbarState2(self, Cookie=defaultNamedNotOptArg,
#                          ToolbarId=defaultNamedNotOptArg,
#                          ToolbarState=defaultNamedNotOptArg):
#         'Gets the state of the toolbar'
#         return self._oleobj_.InvokeTypes(153, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (3, 1)), Cookie
#                                          , ToolbarId, ToolbarState)
#
#     def GetToolbarVisibility(self, Toolbar=defaultNamedNotOptArg):
#         'Get toolbar visibility'
#         return self._oleobj_.InvokeTypes(277, LCID, 1, (11, 0), ((3, 1),),
#                                          Toolbar
#                                          )
#
#     def GetUserPreferenceDoubleValue(self,
#                                      UserPreferenceValue=defaultNamedNotOptArg):
#         "Set the System's User Preference Double Value"
#         return self._oleobj_.InvokeTypes(46, LCID, 1, (5, 0), ((3, 1),),
#                                          UserPreferenceValue
#                                          )
#
#     def GetUserPreferenceIntegerValue(self,
#                                       UserPreferenceValue=defaultNamedNotOptArg):
#         'Get User Preference Integer Value'
#         return self._oleobj_.InvokeTypes(50, LCID, 1, (3, 0), ((3, 1),),
#                                          UserPreferenceValue
#                                          )
#
#     def GetUserPreferenceStringListValue(self,
#                                          UserPreference=defaultNamedNotOptArg):
#         'Get User Preference String List Value'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(66, LCID, 1, (8, 0), ((3, 1),),
#                                          UserPreference
#                                          )
#
#     def GetUserPreferenceStringValue(self,
#                                      UserPreference=defaultNamedNotOptArg):
#         'Get User Preference String Value'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(119, LCID, 1, (8, 0), ((3, 1),),
#                                          UserPreference
#                                          )
#
#     def GetUserPreferenceToggle(self,
#                                 UserPreferenceToggle=defaultNamedNotOptArg):
#         'Get User Preference Toggle'
#         return self._oleobj_.InvokeTypes(44, LCID, 1, (11, 0), ((3, 1),),
#                                          UserPreferenceToggle
#                                          )
#
#     def GetUserProgressBar(self, PProgressBar=pythoncom.Missing):
#         'Get the User Progress Bar'
#         return self._ApplyTypes_(233, 1, (11, 0), ((16393, 2),),
#                                  'GetUserProgressBar', None, PProgressBar
#                                  )
#
#     def GetUserTypeLibReferenceCount(self):
#         'Get the count of User Type Library References'
#         return self._oleobj_.InvokeTypes(204, LCID, 1, (3, 0), (), )
#
#     def GetUserUnit(self, UnitType=defaultNamedNotOptArg):
#         'Get stand alone user units'
#         ret = self._oleobj_.InvokeTypes(142, LCID, 1, (9, 0), ((3, 1),),
#                                         UnitType
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'GetUserUnit', None)
#         return ret
#
#     def HideBubbleTooltip(self):
#         'Hide bubble tip'
#         return self._oleobj_.InvokeTypes(247, LCID, 1, (24, 0), (), )
#
#     def HideToolbar(self, ModuleName=defaultNamedNotOptArg,
#                     ToolbarId=defaultNamedNotOptArg):
#         'Hides a toolbar'
#         return self._oleobj_.InvokeTypes(63, LCID, 1, (11, 0), ((8, 1), (3, 1)),
#                                          ModuleName
#                                          , ToolbarId)
#
#     def HideToolbar2(self, Cookie=defaultNamedNotOptArg,
#                      ToolbarId=defaultNamedNotOptArg):
#         'Hides a toolbar'
#         return self._oleobj_.InvokeTypes(152, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1)), Cookie
#                                          , ToolbarId)
#
#     def HighlightTBButton(self, CmdID=defaultNamedNotOptArg):
#         'HighLight toolbar button'
#         return self._oleobj_.InvokeTypes(176, LCID, 1, (24, 0), ((3, 1),), CmdID
#                                          )
#
#     # Result is of type IModelDoc
#     def IActivateDoc(self, Name=defaultNamedNotOptArg):
#         'Activates a document'
#         ret = self._oleobj_.InvokeTypes(18, LCID, 1, (9, 0), ((8, 1),), Name
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IActivateDoc',
#                            '{83A33D46-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc
#     def IActivateDoc2(self, Name=defaultNamedNotOptArg,
#                       Silent=defaultNamedNotOptArg,
#                       Errors=defaultNamedNotOptArg):
#         'Activates a document'
#         return self._ApplyTypes_(92, 1, (9, 0), ((8, 1), (11, 1), (16387, 3)),
#                                  'IActivateDoc2',
#                                  '{83A33D46-27C5-11CE-BFD4-00400513BB57}', Name
#                                  , Silent, Errors)
#
#     # Result is of type IModelDoc2
#     def IActivateDoc3(self, Name=defaultNamedNotOptArg,
#                       Silent=defaultNamedNotOptArg,
#                       Errors=defaultNamedNotOptArg):
#         'Activates a document'
#         return self._ApplyTypes_(157, 1, (9, 0), ((8, 1), (11, 1), (16387, 3)),
#                                  'IActivateDoc3',
#                                  '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}', Name
#                                  , Silent, Errors)
#
#     def ICopyDocument(self, SourceDoc=defaultNamedNotOptArg,
#                       DestDoc=defaultNamedNotOptArg,
#                       ChildCount=defaultNamedNotOptArg,
#                       FromChildren=defaultNamedNotOptArg
#                       , ToChildren=defaultNamedNotOptArg,
#                       Option=defaultNamedNotOptArg):
#         'Moves a document along with its specified dependents to the destination'
#         return self._oleobj_.InvokeTypes(187, LCID, 1, (3, 0), (
#             (8, 1), (8, 1), (3, 1), (16392, 1), (16392, 1), (3, 1)), SourceDoc
#                                          , DestDoc, ChildCount, FromChildren,
#                                          ToChildren, Option
#                                          )
#
#     # Result is of type IPropertyManagerPage2
#     def ICreatePropertyManagerPage(self, Title=defaultNamedNotOptArg,
#                                    Options=defaultNamedNotOptArg,
#                                    Handler=defaultNamedNotOptArg,
#                                    Errors=defaultNamedNotOptArg):
#         'Create a page for display in the PropertyManager'
#         return self._ApplyTypes_(164, 1, (9, 0),
#                                  ((8, 1), (3, 1), (9, 1), (16387, 3)),
#                                  'ICreatePropertyManagerPage',
#                                  '{B92E624A-0DC3-11D5-AF1E-00C04F603FAF}', Title
#                                  , Options, Handler, Errors)
#
#     # Result is of type IAttributeDef
#     def IDefineAttribute(self, Name=defaultNamedNotOptArg):
#         'Makes an attribute definition'
#         ret = self._oleobj_.InvokeTypes(26, LCID, 1, (9, 0), ((8, 1),), Name
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IDefineAttribute',
#                            '{83A33D67-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     def IEnableStereoDisplay(self, BEnable=defaultNamedNotOptArg):
#         'Enable stereoscopic display view'
#         return self._oleobj_.InvokeTypes(69, LCID, 1, (11, 0), ((11, 1),),
#                                          BEnable
#                                          )
#
#     # Result is of type IFrame
#     def IFrameObject(self):
#         'Gives a handle to the Application Frame'
#         ret = self._oleobj_.InvokeTypes(19, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IFrameObject',
#                            '{83A33D48-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IColorTable
#     def IGetColorTable(self):
#         'Get color table'
#         ret = self._oleobj_.InvokeTypes(140, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetColorTable',
#                            '{83A33DA5-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     def IGetConfigurationNames(self, FilePathName=defaultNamedNotOptArg,
#                                Count=defaultNamedNotOptArg):
#         'Get Configuration Names'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(180, LCID, 1, (8, 0), ((8, 1), (3, 1)),
#                                          FilePathName
#                                          , Count)
#
#     def IGetDocumentDependencies(self, Document=defaultNamedNotOptArg,
#                                  Traverseflag=defaultNamedNotOptArg,
#                                  Searchflag=defaultNamedNotOptArg):
#         'Return names of documents that this document references.'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(71, LCID, 1, (8, 0),
#                                          ((8, 1), (3, 1), (3, 1)), Document
#                                          , Traverseflag, Searchflag)
#
#     def IGetDocumentDependencies2(self, Document=defaultNamedNotOptArg,
#                                   Traverseflag=defaultNamedNotOptArg,
#                                   Searchflag=defaultNamedNotOptArg,
#                                   AddReadOnlyInfo=defaultNamedNotOptArg):
#         'Return names of documents that this document references.'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(105, LCID, 1, (8, 0),
#                                          ((8, 1), (11, 1), (11, 1), (11, 1)),
#                                          Document
#                                          , Traverseflag, Searchflag,
#                                          AddReadOnlyInfo)
#
#     def IGetDocumentDependenciesCount2(self, Document=defaultNamedNotOptArg,
#                                        Traverseflag=defaultNamedNotOptArg,
#                                        Searchflag=defaultNamedNotOptArg,
#                                        AddReadOnlyInfo=defaultNamedNotOptArg):
#         'Returns the size of array needed for a call to IGetDocumentDependencies'
#         return self._oleobj_.InvokeTypes(106, LCID, 1, (3, 0),
#                                          ((8, 1), (11, 1), (11, 1), (11, 1)),
#                                          Document
#                                          , Traverseflag, Searchflag,
#                                          AddReadOnlyInfo)
#
#     # Result is of type IModelDoc2
#     def IGetDocuments(self, NumDocuments=defaultNamedNotOptArg):
#         'Get the open model Documents in SOLIDWORKS'
#         ret = self._oleobj_.InvokeTypes(274, LCID, 1, (9, 0), ((3, 1),),
#                                         NumDocuments
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetDocuments',
#                            '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}')
#         return ret
#
#     # Result is of type IEnvironment
#     def IGetEnvironment(self):
#         'Get the SW environment'
#         ret = self._oleobj_.InvokeTypes(37, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetEnvironment',
#                            '{83A33D78-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc
#     def IGetFirstDocument(self):
#         'Get the first document in the session'
#         ret = self._oleobj_.InvokeTypes(95, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetFirstDocument',
#                            '{83A33D46-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc2
#     def IGetFirstDocument2(self):
#         'Get the first document in the session'
#         ret = self._oleobj_.InvokeTypes(158, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetFirstDocument2',
#                            '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}')
#         return ret
#
#     def IGetMassProperties(self, FilePathName=defaultNamedNotOptArg,
#                            ConfigurationName=defaultNamedNotOptArg,
#                            MPropsData=defaultNamedNotOptArg):
#         'Gets the mass properties from the given document for a given configuration'
#         return self._oleobj_.InvokeTypes(102, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (16389, 1)),
#                                          FilePathName
#                                          , ConfigurationName, MPropsData)
#
#     def IGetMassProperties2(self, FilePathName=defaultNamedNotOptArg,
#                             ConfigurationName=defaultNamedNotOptArg,
#                             MPropsData=defaultNamedNotOptArg,
#                             Accuracy=defaultNamedNotOptArg):
#         'Gets the mass properties from the given document for a given configuration'
#         return self._oleobj_.InvokeTypes(175, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (16389, 1), (3, 1)),
#                                          FilePathName
#                                          , ConfigurationName, MPropsData,
#                                          Accuracy)
#
#     def IGetMaterialDatabases(self, Count=defaultNamedNotOptArg):
#         'Get the list of the material databases.'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(196, LCID, 1, (8, 0), ((3, 1),), Count
#                                          )
#
#     # Result is of type IMathUtility
#     def IGetMathUtility(self):
#         'Gets Math Utility Interface'
#         ret = self._oleobj_.InvokeTypes(133, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetMathUtility',
#                            '{F7D97F80-162E-11D4-AEAB-00C04FA0AC51}')
#         return ret
#
#     # Result is of type IModeler
#     def IGetModeler(self):
#         'Get the Geometry Modeler'
#         ret = self._oleobj_.InvokeTypes(35, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetModeler',
#                            '{83A33D73-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc
#     def IGetOpenDocumentByName(self, DocumentName=defaultNamedNotOptArg):
#         'Gets the open document with the given name'
#         ret = self._oleobj_.InvokeTypes(123, LCID, 1, (9, 0), ((8, 1),),
#                                         DocumentName
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetOpenDocumentByName',
#                            '{83A33D46-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc2
#     def IGetOpenDocumentByName2(self, DocumentName=defaultNamedNotOptArg):
#         'Gets the open document with the given name'
#         ret = self._oleobj_.InvokeTypes(160, LCID, 1, (9, 0), ((8, 1),),
#                                         DocumentName
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetOpenDocumentByName2',
#                            '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}')
#         return ret
#
#     # Result is of type IRayTraceRenderer
#     def IGetRayTraceRenderer(self, RendererType=defaultNamedNotOptArg):
#         'Get the ray trace renderer'
#         ret = self._oleobj_.InvokeTypes(297, LCID, 1, (9, 0), ((3, 1),),
#                                         RendererType
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetRayTraceRenderer',
#                            '{D9920A70-9CA9-47F9-BC83-264DECA00C8A}')
#         return ret
#
#     def IGetSelectionFilters(self):
#         'Get all selection filters with status=ON'
#         return self._oleobj_.InvokeTypes(115, LCID, 1, (3, 0), (), )
#
#     def IGetSelectionFiltersCount(self):
#         'Get the count of selection filters with status=ON'
#         return self._oleobj_.InvokeTypes(114, LCID, 1, (3, 0), (), )
#
#     def IGetTemplateSizes(self, FileName=defaultNamedNotOptArg,
#                           PaperSize=pythoncom.Missing, Width=pythoncom.Missing,
#                           Height=pythoncom.Missing):
#         'Get the sheet properties from a template document'
#         return self._ApplyTypes_(138, 1, (11, 0),
#                                  ((8, 1), (16387, 2), (16389, 2), (16389, 2)),
#                                  'IGetTemplateSizes', None, FileName
#                                  , PaperSize, Width, Height)
#
#     def IGetUserTypeLibReferences(self, NCount=defaultNamedNotOptArg):
#         'Get the User Type Library References'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(206, LCID, 1, (8, 0), ((3, 1),), NCount
#                                          )
#
#     # Result is of type IUserUnit
#     def IGetUserUnit(self, UnitType=defaultNamedNotOptArg):
#         'Get stand alone user units'
#         ret = self._oleobj_.InvokeTypes(143, LCID, 1, (9, 0), ((3, 1),),
#                                         UnitType
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'IGetUserUnit',
#                            '{82071121-8B32-4F51-8983-9304756503E7}')
#         return ret
#
#     def IGetVersionHistoryCount(self, FileName=defaultNamedNotOptArg):
#         'Gets the number of version history strings'
#         return self._oleobj_.InvokeTypes(83, LCID, 1, (3, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def IMoveDocument(self, SourceDoc=defaultNamedNotOptArg,
#                       DestDoc=defaultNamedNotOptArg,
#                       ChildCount=defaultNamedNotOptArg,
#                       FromChildren=defaultNamedNotOptArg
#                       , ToChildren=defaultNamedNotOptArg,
#                       Option=defaultNamedNotOptArg):
#         'Moves a document along with its specified dependents to the destination'
#         return self._oleobj_.InvokeTypes(186, LCID, 1, (3, 0), (
#             (8, 1), (8, 1), (3, 1), (16392, 1), (16392, 1), (3, 1)), SourceDoc
#                                          , DestDoc, ChildCount, FromChildren,
#                                          ToChildren, Option
#                                          )
#
#     # Result is of type IAssemblyDoc
#     def INewAssembly(self):
#         'Creates a new assembly document'
#         ret = self._oleobj_.InvokeTypes(21, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'INewAssembly',
#                            '{83A33D35-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc
#     def INewDocument(self, TemplateName=defaultNamedNotOptArg,
#                      PaperSize=defaultNamedNotOptArg,
#                      Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
#         'Creates a new document based on the template name'
#         ret = self._oleobj_.InvokeTypes(112, LCID, 1, (9, 0),
#                                         ((8, 1), (3, 1), (5, 1), (5, 1)),
#                                         TemplateName
#                                         , PaperSize, Width, Height)
#         if ret is not None:
#             ret = Dispatch(ret, 'INewDocument',
#                            '{83A33D46-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc2
#     def INewDocument2(self, TemplateName=defaultNamedNotOptArg,
#                       PaperSize=defaultNamedNotOptArg,
#                       Width=defaultNamedNotOptArg,
#                       Height=defaultNamedNotOptArg):
#         'Creates a new document based on the template name'
#         ret = self._oleobj_.InvokeTypes(159, LCID, 1, (9, 0),
#                                         ((8, 1), (3, 1), (5, 1), (5, 1)),
#                                         TemplateName
#                                         , PaperSize, Width, Height)
#         if ret is not None:
#             ret = Dispatch(ret, 'INewDocument2',
#                            '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}')
#         return ret
#
#     # Result is of type IDrawingDoc
#     def INewDrawing(self, TemplateToUse=defaultNamedNotOptArg):
#         'Creates a new drawing document'
#         ret = self._oleobj_.InvokeTypes(22, LCID, 1, (9, 0), ((3, 1),),
#                                         TemplateToUse
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'INewDrawing',
#                            '{83A33D33-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IDrawingDoc
#     def INewDrawing2(self, TemplateToUse=defaultNamedNotOptArg,
#                      TemplateName=defaultNamedNotOptArg,
#                      PaperSize=defaultNamedNotOptArg,
#                      Width=defaultNamedNotOptArg
#                      , Height=defaultNamedNotOptArg):
#         'Creates a new drawing document'
#         ret = self._oleobj_.InvokeTypes(39, LCID, 1, (9, 0), (
#             (3, 1), (8, 1), (3, 1), (5, 1), (5, 1)), TemplateToUse
#                                         , TemplateName, PaperSize, Width,
#                                         Height)
#         if ret is not None:
#             ret = Dispatch(ret, 'INewDrawing2',
#                            '{83A33D33-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IPartDoc
#     def INewPart(self):
#         'Creates a new part document'
#         ret = self._oleobj_.InvokeTypes(20, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'INewPart',
#                            '{83A33D32-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc
#     def IOpenDoc(self, Name=defaultNamedNotOptArg, Type=defaultNamedNotOptArg):
#         'Opens an existing document'
#         ret = self._oleobj_.InvokeTypes(17, LCID, 1, (9, 0), ((8, 1), (3, 1)),
#                                         Name
#                                         , Type)
#         if ret is not None:
#             ret = Dispatch(ret, 'IOpenDoc',
#                            '{83A33D46-27C5-11CE-BFD4-00400513BB57}')
#         return ret
#
#     # Result is of type IModelDoc
#     def IOpenDoc2(self, FileName=defaultNamedNotOptArg,
#                   Type=defaultNamedNotOptArg, ReadOnly=defaultNamedNotOptArg,
#                   ViewOnly=defaultNamedNotOptArg
#                   , Silent=defaultNamedNotOptArg, Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(100, 1, (9, 0), (
#             (8, 1), (3, 1), (11, 1), (11, 1), (11, 1), (16387, 3)), 'IOpenDoc2',
#                                  '{83A33D46-27C5-11CE-BFD4-00400513BB57}',
#                                  FileName
#                                  , Type, ReadOnly, ViewOnly, Silent, Errors
#                                  )
#
#     # Result is of type IModelDoc
#     def IOpenDoc3(self, FileName=defaultNamedNotOptArg,
#                   Type=defaultNamedNotOptArg, ReadOnly=defaultNamedNotOptArg,
#                   ViewOnly=defaultNamedNotOptArg
#                   , RapidDraft=defaultNamedNotOptArg,
#                   Silent=defaultNamedNotOptArg, Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(127, 1, (9, 0), (
#             (8, 1), (3, 1), (11, 1), (11, 1), (11, 1), (11, 1), (16387, 3)),
#                                  'IOpenDoc3',
#                                  '{83A33D46-27C5-11CE-BFD4-00400513BB57}',
#                                  FileName
#                                  , Type, ReadOnly, ViewOnly, RapidDraft, Silent
#                                  , Errors)
#
#     # Result is of type IModelDoc
#     def IOpenDoc4(self, FileName=defaultNamedNotOptArg,
#                   Type=defaultNamedNotOptArg, Options=defaultNamedNotOptArg,
#                   Configuration=defaultNamedNotOptArg
#                   , Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(135, 1, (9, 0),
#                                  ((8, 1), (3, 1), (3, 1), (8, 1), (16387, 3)),
#                                  'IOpenDoc4',
#                                  '{83A33D46-27C5-11CE-BFD4-00400513BB57}',
#                                  FileName
#                                  , Type, Options, Configuration, Errors)
#
#     # Result is of type IModelDoc2
#     def IOpenDoc5(self, FileName=defaultNamedNotOptArg,
#                   Type=defaultNamedNotOptArg, Options=defaultNamedNotOptArg,
#                   Configuration=defaultNamedNotOptArg
#                   , Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(161, 1, (9, 0),
#                                  ((8, 1), (3, 1), (3, 1), (8, 1), (16387, 3)),
#                                  'IOpenDoc5',
#                                  '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}',
#                                  FileName
#                                  , Type, Options, Configuration, Errors)
#
#     # Result is of type IModelDoc
#     def IOpenDocSilent(self, FileName=defaultNamedNotOptArg,
#                        Type=defaultNamedNotOptArg,
#                        Errors=defaultNamedNotOptArg):
#         'Opens a document with suppresion of various error & warning dialogs'
#         return self._ApplyTypes_(74, 1, (9, 0), ((8, 1), (3, 1), (16387, 3)),
#                                  'IOpenDocSilent',
#                                  '{83A33D46-27C5-11CE-BFD4-00400513BB57}',
#                                  FileName
#                                  , Type, Errors)
#
#     def IRemoveUserTypeLibReferences(self, NCount=defaultNamedNotOptArg,
#                                      BstrTlbRef=defaultNamedNotOptArg):
#         'Remove User Type Library References'
#         return self._oleobj_.InvokeTypes(209, LCID, 1, (11, 0),
#                                          ((3, 1), (16392, 1)), NCount
#                                          , BstrTlbRef)
#
#     def ISetSelectionFilters(self, Count=defaultNamedNotOptArg,
#                              SelType=defaultNamedNotOptArg,
#                              State=defaultNamedNotOptArg):
#         'Set status for multiple selection filters'
#         return self._oleobj_.InvokeTypes(116, LCID, 1, (24, 0),
#                                          ((3, 1), (16387, 1), (11, 1)), Count
#                                          , SelType, State)
#
#     def ISetUserTypeLibReferences(self, NCount=defaultNamedNotOptArg,
#                                   BstrTlbRef=defaultNamedNotOptArg):
#         'Set the User Type Library References'
#         return self._oleobj_.InvokeTypes(207, LCID, 1, (24, 0),
#                                          ((3, 1), (16392, 1)), NCount
#                                          , BstrTlbRef)
#
#     def IVersionHistory(self, FileName=defaultNamedNotOptArg):
#         'Get the version history of the given file'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(82, LCID, 1, (8, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def ImportHoleWizardItem(self, StdToImport=defaultNamedNotOptArg,
#                              DestinationFilePath=defaultNamedNotOptArg,
#                              ReplaceData=defaultNamedNotOptArg,
#                              ErrorFile=defaultNamedNotOptArg):
#         'Import HoleWizard Item'
#         return self._oleobj_.InvokeTypes(322, LCID, 1, (3, 0),
#                                          ((8, 1), (8, 1), (11, 1), (11, 1)),
#                                          StdToImport
#                                          , DestinationFilePath, ReplaceData,
#                                          ErrorFile)
#
#     def ImportToolboxItem(self, StdToImport=defaultNamedNotOptArg,
#                           DestinationFilePath=defaultNamedNotOptArg):
#         'Import Toolbox Item'
#         return self._oleobj_.InvokeTypes(324, LCID, 1, (3, 0), ((8, 1), (8, 1)),
#                                          StdToImport
#                                          , DestinationFilePath)
#
#     def InstallQuickTipGuide(self, PInterface=defaultNamedNotOptArg):
#         'Installs a quickTip guide'
#         return self._oleobj_.InvokeTypes(199, LCID, 1, (24, 0), ((9, 1),),
#                                          PInterface
#                                          )
#
#     def IsBackgroundProcessingCompleted(self, FilePath=defaultNamedNotOptArg):
#         'Get status whether background processing is completed or not'
#         return self._oleobj_.InvokeTypes(288, LCID, 1, (11, 0), ((8, 1),),
#                                          FilePath
#                                          )
#
#     def IsCommandEnabled(self, CommandID=defaultNamedNotOptArg):
#         'Is Command enabled'
#         return self._oleobj_.InvokeTypes(271, LCID, 1, (11, 0), ((3, 1),),
#                                          CommandID
#                                          )
#
#     def IsRapidDraft(self, FileName=defaultNamedNotOptArg):
#         'Is a given drawing file in Rapid Draft format'
#         return self._oleobj_.InvokeTypes(136, LCID, 1, (11, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def IsSame(self, Object1=defaultNamedNotOptArg,
#                Object2=defaultNamedNotOptArg):
#         'Compares two object to determine if they are the same'
#         return self._oleobj_.InvokeTypes(283, LCID, 1, (3, 0), ((9, 1), (9, 1)),
#                                          Object1
#                                          , Object2)
#
#     def LoadAddIn(self, FileName=defaultNamedNotOptArg):
#         'Load an Add-In DLL'
#         return self._oleobj_.InvokeTypes(78, LCID, 1, (3, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def LoadFile(self, FileName=defaultNamedNotOptArg):
#         'Loads a foreign file'
#         return self._oleobj_.InvokeTypes(13, LCID, 1, (11, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def LoadFile2(self, FileName=defaultNamedNotOptArg,
#                   ArgString=defaultNamedNotOptArg):
#         'Loads a foreign file'
#         return self._oleobj_.InvokeTypes(49, LCID, 1, (11, 0), ((8, 1), (8, 1)),
#                                          FileName
#                                          , ArgString)
#
#     def LoadFile3(self, FileName=defaultNamedNotOptArg,
#                   ArgString=defaultNamedNotOptArg,
#                   ImportData=defaultNamedNotOptArg):
#         'Load a foreign file'
#         return self._oleobj_.InvokeTypes(218, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (9, 1)), FileName
#                                          , ArgString, ImportData)
#
#     # Result is of type IModelDoc2
#     def LoadFile4(self, FileName=defaultNamedNotOptArg,
#                   ArgString=defaultNamedNotOptArg,
#                   ImportData=defaultNamedNotOptArg,
#                   Errors=defaultNamedNotOptArg):
#         'Loads a foreign file'
#         return self._ApplyTypes_(227, 1, (9, 0),
#                                  ((8, 1), (8, 1), (9, 1), (16387, 3)),
#                                  'LoadFile4',
#                                  '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}',
#                                  FileName
#                                  , ArgString, ImportData, Errors)
#
#     def MoveDocument(self, SourceDoc=defaultNamedNotOptArg,
#                      DestDoc=defaultNamedNotOptArg,
#                      FromChildren=defaultNamedNotOptArg,
#                      ToChildren=defaultNamedNotOptArg
#                      , Option=defaultNamedNotOptArg):
#         'Moves a document along with its specified dependents to the destination'
#         return self._oleobj_.InvokeTypes(184, LCID, 1, (3, 0), (
#             (8, 1), (8, 1), (12, 1), (12, 1), (3, 1)), SourceDoc
#                                          , DestDoc, FromChildren, ToChildren,
#                                          Option)
#
#     def NewAssembly(self):
#         'Creates a new assembly document'
#         ret = self._oleobj_.InvokeTypes(9, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'NewAssembly', None)
#         return ret
#
#     def NewDocument(self, TemplateName=defaultNamedNotOptArg,
#                     PaperSize=defaultNamedNotOptArg,
#                     Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
#         'Creates a new document based on the template name'
#         ret = self._oleobj_.InvokeTypes(111, LCID, 1, (9, 0),
#                                         ((8, 1), (3, 1), (5, 1), (5, 1)),
#                                         TemplateName
#                                         , PaperSize, Width, Height)
#         if ret is not None:
#             ret = Dispatch(ret, 'NewDocument', None)
#         return ret
#
#     def NewDrawing(self, TemplateToUse=defaultNamedNotOptArg):
#         'Creates a new drawing document'
#         ret = self._oleobj_.InvokeTypes(10, LCID, 1, (9, 0), ((3, 1),),
#                                         TemplateToUse
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'NewDrawing', None)
#         return ret
#
#     def NewDrawing2(self, TemplateToUse=defaultNamedNotOptArg,
#                     TemplateName=defaultNamedNotOptArg,
#                     PaperSize=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
#                     , Height=defaultNamedNotOptArg):
#         'Creates a new drawing document'
#         ret = self._oleobj_.InvokeTypes(38, LCID, 1, (9, 0), (
#             (3, 1), (8, 1), (3, 1), (5, 1), (5, 1)), TemplateToUse
#                                         , TemplateName, PaperSize, Width,
#                                         Height)
#         if ret is not None:
#             ret = Dispatch(ret, 'NewDrawing2', None)
#         return ret
#
#     def NewPart(self):
#         'Creates a new part document'
#         ret = self._oleobj_.InvokeTypes(8, LCID, 1, (9, 0), (), )
#         if ret is not None:
#             ret = Dispatch(ret, 'NewPart', None)
#         return ret
#
#     def OpenDoc(self, Name=defaultNamedNotOptArg, Type=defaultNamedNotOptArg):
#         'Opens an existing document'
#         ret = self._oleobj_.InvokeTypes(2, LCID, 1, (9, 0), ((8, 1), (3, 1)),
#                                         Name
#                                         , Type)
#         if ret is not None:
#             ret = Dispatch(ret, 'OpenDoc', None)
#         return ret
#
#     def OpenDoc2(self, FileName=defaultNamedNotOptArg,
#                  Type=defaultNamedNotOptArg, ReadOnly=defaultNamedNotOptArg,
#                  ViewOnly=defaultNamedNotOptArg
#                  , Silent=defaultNamedNotOptArg, Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(99, 1, (9, 0), (
#             (8, 1), (3, 1), (11, 1), (11, 1), (11, 1), (16387, 3)), 'OpenDoc2',
#                                  None, FileName
#                                  , Type, ReadOnly, ViewOnly, Silent, Errors
#                                  )
#
#     def OpenDoc3(self, FileName=defaultNamedNotOptArg,
#                  Type=defaultNamedNotOptArg, ReadOnly=defaultNamedNotOptArg,
#                  ViewOnly=defaultNamedNotOptArg
#                  , RapidDraft=defaultNamedNotOptArg,
#                  Silent=defaultNamedNotOptArg, Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(126, 1, (9, 0), (
#             (8, 1), (3, 1), (11, 1), (11, 1), (11, 1), (11, 1), (16387, 3)),
#                                  'OpenDoc3', None, FileName
#                                  , Type, ReadOnly, ViewOnly, RapidDraft, Silent
#                                  , Errors)
#
#     def OpenDoc4(self, FileName=defaultNamedNotOptArg,
#                  Type=defaultNamedNotOptArg, Options=defaultNamedNotOptArg,
#                  Configuration=defaultNamedNotOptArg
#                  , Errors=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(134, 1, (9, 0),
#                                  ((8, 1), (3, 1), (3, 1), (8, 1), (16387, 3)),
#                                  'OpenDoc4', None, FileName
#                                  , Type, Options, Configuration, Errors)
#
#     # Result is of type IModelDoc2
#     def OpenDoc6(self, FileName=defaultNamedNotOptArg,
#                  Type=defaultNamedNotOptArg, Options=defaultNamedNotOptArg,
#                  Configuration=defaultNamedNotOptArg
#                  , Errors=defaultNamedNotOptArg,
#                  Warnings=defaultNamedNotOptArg):
#         'Opens an existing document'
#         return self._ApplyTypes_(167, 1, (9, 0), (
#             (8, 1), (3, 1), (3, 1), (8, 1), (16387, 3), (16387, 3)), 'OpenDoc6',
#                                  '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}',
#                                  FileName
#                                  , Type, Options, Configuration, Errors,
#                                  Warnings
#                                  )
#
#     # Result is of type IModelDoc2
#     def OpenDoc7(self, Specification=defaultNamedNotOptArg):
#         'Opens an existing document'
#         ret = self._oleobj_.InvokeTypes(249, LCID, 1, (9, 0), ((9, 1),),
#                                         Specification
#                                         )
#         if ret is not None:
#             ret = Dispatch(ret, 'OpenDoc7',
#                            '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}')
#         return ret
#
#     def OpenDocSilent(self, FileName=defaultNamedNotOptArg,
#                       Type=defaultNamedNotOptArg, Errors=defaultNamedNotOptArg):
#         'Opens a document with suppresion of various error & warning dialogs'
#         return self._ApplyTypes_(73, 1, (9, 0), ((8, 1), (3, 1), (16387, 3)),
#                                  'OpenDocSilent', None, FileName
#                                  , Type, Errors)
#
#     def OpenModelConfiguration(self, PathName=defaultNamedNotOptArg,
#                                ConfigName=defaultNamedNotOptArg):
#         'Open model in supplied configuration'
#         ret = self._oleobj_.InvokeTypes(129, LCID, 1, (9, 0), ((8, 1), (8, 1)),
#                                         PathName
#                                         , ConfigName)
#         if ret is not None:
#             ret = Dispatch(ret, 'OpenModelConfiguration', None)
#         return ret
#
#     def PasteAppearance(self, Object=defaultNamedNotOptArg,
#                         AppearanceTarget=defaultNamedNotOptArg):
#         'Paste appearance to input object or selected object or appearance target'
#         return self._oleobj_.InvokeTypes(308, LCID, 1, (11, 0),
#                                          ((9, 1), (3, 1)), Object
#                                          , AppearanceTarget)
#
#     def PostMessageToApplication(self, Cookie=defaultNamedNotOptArg,
#                                  UserData=defaultNamedNotOptArg):
#         'Post a message to the sending application'
#         return self._oleobj_.InvokeTypes(291, LCID, 1, (24, 0),
#                                          ((3, 1), (3, 1)), Cookie
#                                          , UserData)
#
#     def PostMessageToApplicationx64(self, Cookie=defaultNamedNotOptArg,
#                                     UserData=defaultNamedNotOptArg):
#         'Post a message to the sending application'
#         return self._oleobj_.InvokeTypes(301, LCID, 1, (24, 0),
#                                          ((3, 1), (20, 1)), Cookie
#                                          , UserData)
#
#     def PreSelectDwgTemplateSize(self, TemplateToUse=defaultNamedNotOptArg,
#                                  TemplateName=defaultNamedNotOptArg):
#         'PreSelects drawing template size'
#         return self._oleobj_.InvokeTypes(23, LCID, 1, (24, 0), ((3, 1), (8, 1)),
#                                          TemplateToUse
#                                          , TemplateName)
#
#     def PresetNewDrawingParameters(self, DrawingTemplate=defaultNamedNotOptArg,
#                                    ShowTemplate=defaultNamedNotOptArg,
#                                    Width=defaultNamedNotOptArg,
#                                    Height=defaultNamedNotOptArg):
#         'Presets drawing template and sheet size parameters to avoid showing the drawing template dialog in the UI'
#         return self._oleobj_.InvokeTypes(242, LCID, 1, (11, 0),
#                                          ((8, 1), (11, 1), (5, 1), (5, 1)),
#                                          DrawingTemplate
#                                          , ShowTemplate, Width, Height)
#
#     def PreviewDoc(self, HWnd=defaultNamedNotOptArg,
#                    FullName=defaultNamedNotOptArg):
#         'Display a preview of a given document to the given window'
#         return self._oleobj_.InvokeTypes(41, LCID, 1, (11, 0),
#                                          ((16387, 1), (8, 1)), HWnd
#                                          , FullName)
#
#     def PreviewDocx64(self, HWnd=defaultNamedNotOptArg,
#                       FullName=defaultNamedNotOptArg):
#         'Display a preview of a given document to the given window'
#         return self._oleobj_.InvokeTypes(231, LCID, 1, (11, 0),
#                                          ((16404, 1), (8, 1)), HWnd
#                                          , FullName)
#
#     def QuitDoc(self, Name=defaultNamedNotOptArg):
#         'Close the named document without saving changes'
#         return self._oleobj_.InvokeTypes(33, LCID, 1, (24, 0), ((8, 1),), Name
#                                          )
#
#     def RecordLine(self, Text=defaultNamedNotOptArg):
#         'Add a specified line of text to the journal and macro recording files'
#         return self._oleobj_.InvokeTypes(80, LCID, 1, (11, 0), ((8, 1),), Text
#                                          )
#
#     def RecordLineCSharp(self, StringLine=defaultNamedNotOptArg):
#         'Record line in CSharp format'
#         return self._oleobj_.InvokeTypes(299, LCID, 1, (11, 0), ((8, 1),),
#                                          StringLine
#                                          )
#
#     def RecordLineVBnet(self, StringLine=defaultNamedNotOptArg):
#         'Record line in Vb.net format'
#         return self._oleobj_.InvokeTypes(298, LCID, 1, (11, 0), ((8, 1),),
#                                          StringLine
#                                          )
#
#     def RefreshQuickTipWindow(self):
#         'Refreshes the quickTip guide window'
#         return self._oleobj_.InvokeTypes(201, LCID, 1, (24, 0), (), )
#
#     def RefreshTaskpaneContent(self):
#         'Refresh Content Manager Tree and View'
#         return self._oleobj_.InvokeTypes(241, LCID, 1, (24, 0), (), )
#
#     def RegisterThirdPartyPopupMenu(self):
#         'Register third party popup menu'
#         return self._oleobj_.InvokeTypes(280, LCID, 1, (3, 0), (), )
#
#     def RegisterTrackingDefinition(self, Name=defaultNamedNotOptArg):
#         'Register an Tracking Definition'
#         return self._oleobj_.InvokeTypes(263, LCID, 1, (3, 0), ((8, 1),), Name
#                                          )
#
#     def RemoveCallback(self, Cookie=defaultNamedNotOptArg):
#         'Unregister a general perpose callback handler'
#         return self._oleobj_.InvokeTypes(223, LCID, 1, (24, 0), ((3, 1),),
#                                          Cookie
#                                          )
#
#     def RemoveFileOpenItem(self, CallbackFcnAndModule=defaultNamedNotOptArg,
#                            Description=defaultNamedNotOptArg):
#         'Removes an item from the Open drop down that was added with AddFileOpenItem'
#         return self._oleobj_.InvokeTypes(54, LCID, 1, (11, 0), ((8, 1), (8, 1)),
#                                          CallbackFcnAndModule
#                                          , Description)
#
#     def RemoveFileOpenItem2(self, Cookie=defaultNamedNotOptArg,
#                             MethodName=defaultNamedNotOptArg,
#                             Description=defaultNamedNotOptArg,
#                             Extension=defaultNamedNotOptArg):
#         'Removes an item from the Open drop down that was added with AddFileOpenItem'
#         return self._oleobj_.InvokeTypes(169, LCID, 1, (11, 0),
#                                          ((3, 1), (8, 1), (8, 1), (8, 1)),
#                                          Cookie
#                                          , MethodName, Description, Extension)
#
#     def RemoveFileSaveAsItem(self, CallbackFcnAndModule=defaultNamedNotOptArg,
#                              Description=defaultNamedNotOptArg,
#                              Type=defaultNamedNotOptArg):
#         'Removes an item from the Save As drop down which was added with AddFileSaveAsItem'
#         return self._oleobj_.InvokeTypes(55, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (3, 1)),
#                                          CallbackFcnAndModule
#                                          , Description, Type)
#
#     def RemoveFileSaveAsItem2(self, Cookie=defaultNamedNotOptArg,
#                               MethodName=defaultNamedNotOptArg,
#                               Description=defaultNamedNotOptArg,
#                               Extension=defaultNamedNotOptArg
#                               , DocumentType=defaultNamedNotOptArg):
#         'Removes an item from the Save As drop down which was added with AddFileSaveAsItem'
#         return self._oleobj_.InvokeTypes(171, LCID, 1, (11, 0), (
#             (3, 1), (8, 1), (8, 1), (8, 1), (3, 1)), Cookie
#                                          , MethodName, Description, Extension,
#                                          DocumentType)
#
#     def RemoveFromMenu(self, CommandID=defaultNamedNotOptArg,
#                        DocumentType=defaultNamedNotOptArg,
#                        Option=defaultNamedNotOptArg,
#                        RemoveParentMenu=defaultNamedNotOptArg):
#         'Remove menu item by command ID'
#         return self._oleobj_.InvokeTypes(238, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (3, 1), (11, 1)),
#                                          CommandID
#                                          , DocumentType, Option,
#                                          RemoveParentMenu)
#
#     def RemoveFromPopupMenu(self, CommandID=defaultNamedNotOptArg,
#                             DocumentType=defaultNamedNotOptArg,
#                             SelectionType=defaultNamedNotOptArg,
#                             RemoveParentMenu=defaultNamedNotOptArg):
#         'Remove context menu item by command ID'
#         return self._oleobj_.InvokeTypes(239, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (3, 1), (11, 1)),
#                                          CommandID
#                                          , DocumentType, SelectionType,
#                                          RemoveParentMenu)
#
#     def RemoveItemFromThirdPartyPopupMenu(self,
#                                           RegisterId=defaultNamedNotOptArg,
#                                           DocType=defaultNamedNotOptArg,
#                                           Item=defaultNamedNotOptArg,
#                                           IconIndex=defaultNamedNotOptArg):
#         'Add item to third party popup menu'
#         return self._oleobj_.InvokeTypes(290, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (8, 1), (3, 1)),
#                                          RegisterId
#                                          , DocType, Item, IconIndex)
#
#     def RemoveMenu(self, DocType=defaultNamedNotOptArg,
#                    MenuItemString=defaultNamedNotOptArg,
#                    CallbackFcnAndModule=defaultNamedNotOptArg):
#         'Removes Menu'
#         return self._oleobj_.InvokeTypes(53, LCID, 1, (11, 0),
#                                          ((3, 1), (8, 1), (8, 1)), DocType
#                                          , MenuItemString, CallbackFcnAndModule)
#
#     def RemoveMenuPopupItem(self, DocType=defaultNamedNotOptArg,
#                             SelectType=defaultNamedNotOptArg,
#                             Item=defaultNamedNotOptArg,
#                             CallbackFcnAndModule=defaultNamedNotOptArg
#                             , CustomNames=defaultNamedNotOptArg,
#                             Unused=defaultNamedNotOptArg):
#         'Removes Popup Menu Item'
#         return self._oleobj_.InvokeTypes(52, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (3, 1)), DocType
#                                          , SelectType, Item,
#                                          CallbackFcnAndModule, CustomNames,
#                                          Unused
#                                          )
#
#     def RemoveMenuPopupItem2(self, DocumentType=defaultNamedNotOptArg,
#                              Cookie=defaultNamedNotOptArg,
#                              SelectType=defaultNamedNotOptArg,
#                              PopupItemName=defaultNamedNotOptArg
#                              , MenuCallback=defaultNamedNotOptArg,
#                              MenuEnableMethod=defaultNamedNotOptArg,
#                              HintString=defaultNamedNotOptArg,
#                              CustomNames=defaultNamedNotOptArg):
#         'Removes Popup Menu Item'
#         return self._oleobj_.InvokeTypes(173, LCID, 1, (11, 0), (
#             (3, 1), (3, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1)),
#                                          DocumentType
#                                          , Cookie, SelectType, PopupItemName,
#                                          MenuCallback, MenuEnableMethod
#                                          , HintString, CustomNames)
#
#     def RemoveToolbar(self, Module=defaultNamedNotOptArg,
#                       ToolbarId=defaultNamedNotOptArg):
#         'Remove a toolbar'
#         return self._oleobj_.InvokeTypes(64, LCID, 1, (11, 0), ((8, 1), (3, 1)),
#                                          Module
#                                          , ToolbarId)
#
#     def RemoveToolbar2(self, Cookie=defaultNamedNotOptArg,
#                        ToolbarId=defaultNamedNotOptArg):
#         'Remove a toolbar'
#         return self._oleobj_.InvokeTypes(149, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1)), Cookie
#                                          , ToolbarId)
#
#     def RemoveUserMenu(self, DocType=defaultNamedNotOptArg,
#                        MenuIdIn=defaultNamedNotOptArg,
#                        ModuleName=defaultNamedNotOptArg):
#         'Remove a menu item'
#         return self._oleobj_.InvokeTypes(59, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (8, 1)), DocType
#                                          , MenuIdIn, ModuleName)
#
#     def RemoveUserTypeLibReferences(self, VTlbRef=defaultNamedNotOptArg):
#         'Remove User Type Library References'
#         return self._oleobj_.InvokeTypes(208, LCID, 1, (11, 0), ((12, 1),),
#                                          VTlbRef
#                                          )
#
#     def ReplaceReferencedDocument(self,
#                                   ReferencingDocument=defaultNamedNotOptArg,
#                                   ReferencedDocument=defaultNamedNotOptArg,
#                                   NewReference=defaultNamedNotOptArg):
#         'In the specified document, replace a referenced document with another document.'
#         return self._oleobj_.InvokeTypes(56, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (8, 1)),
#                                          ReferencingDocument
#                                          , ReferencedDocument, NewReference)
#
#     def ResetPresetDrawingParameters(self):
#         'Resets new drawing parameters set in PresetNewDrawingParameters'
#         return self._oleobj_.InvokeTypes(243, LCID, 1, (24, 0), (), )
#
#     def ResetUntitledCount(self, PartValue=defaultNamedNotOptArg,
#                            AssemValue=defaultNamedNotOptArg,
#                            DrawingValue=defaultNamedNotOptArg):
#         'Resets the index for new untitled documents'
#         return self._oleobj_.InvokeTypes(276, LCID, 1, (3, 0),
#                                          ((3, 1), (3, 1), (3, 1)), PartValue
#                                          , AssemValue, DrawingValue)
#
#     def RestoreSettings(self, FileName=defaultNamedNotOptArg,
#                         SystemOptions=defaultNamedNotOptArg,
#                         ToolbarLayout=defaultNamedNotOptArg,
#                         KeyboardShortcuts=defaultNamedNotOptArg
#                         , MouseGestures=defaultNamedNotOptArg,
#                         MenuCustomization=defaultNamedNotOptArg,
#                         SavedViews=defaultNamedNotOptArg,
#                         CreateBackup=defaultNamedNotOptArg):
#         'Restore Settings'
#         return self._oleobj_.InvokeTypes(320, LCID, 1, (3, 0), (
#             (8, 1), (11, 1), (11, 1), (11, 1), (11, 1), (11, 1), (11, 1),
#             (11, 1)),
#                                          FileName
#                                          , SystemOptions, ToolbarLayout,
#                                          KeyboardShortcuts, MouseGestures,
#                                          MenuCustomization
#                                          , SavedViews, CreateBackup)
#
#     def ResumeSkinning(self):
#         'Resume Skinning , skinning a window'
#         return self._oleobj_.InvokeTypes(251, LCID, 1, (11, 0), (), )
#
#     def RevisionNumber(self):
#         'Returns the revision number of the application'
#         # Result is a Unicode object
#         return self._oleobj_.InvokeTypes(12, LCID, 1, (8, 0), (), )
#
#     def RunAttachedMacro(self, FileName=defaultNamedNotOptArg,
#                          ModuleName=defaultNamedNotOptArg,
#                          ProcedureName=defaultNamedNotOptArg):
#         'Run attached design binder macro'
#         return self._oleobj_.InvokeTypes(269, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (8, 1)), FileName
#                                          , ModuleName, ProcedureName)
#
#     def RunCommand(self, CommandID=defaultNamedNotOptArg,
#                    NewTitle=defaultNamedNotOptArg):
#         'Run a command with specified title'
#         return self._oleobj_.InvokeTypes(245, LCID, 1, (11, 0),
#                                          ((3, 1), (8, 1)), CommandID
#                                          , NewTitle)
#
#     def RunJournalCmd(self, Cmd=defaultNamedNotOptArg):
#         'Run Serialized Journaling Command'
#         return self._oleobj_.InvokeTypes(285, LCID, 1, (11, 0), ((8, 1),), Cmd
#                                          )
#
#     def RunMacro(self, FilePathName=defaultNamedNotOptArg,
#                  ModuleName=defaultNamedNotOptArg,
#                  ProcedureName=defaultNamedNotOptArg):
#         'Run macro from a project file'
#         return self._oleobj_.InvokeTypes(177, LCID, 1, (11, 0),
#                                          ((8, 1), (8, 1), (8, 1)), FilePathName
#                                          , ModuleName, ProcedureName)
#
#     def RunMacro2(self, FilePathName=defaultNamedNotOptArg,
#                   ModuleName=defaultNamedNotOptArg,
#                   ProcedureName=defaultNamedNotOptArg,
#                   Options=defaultNamedNotOptArg
#                   , Error=pythoncom.Missing):
#         'Run macro from a project file'
#         return self._ApplyTypes_(270, 1, (11, 0),
#                                  ((8, 1), (8, 1), (8, 1), (3, 1), (16387, 2)),
#                                  'RunMacro2', None, FilePathName
#                                  , ModuleName, ProcedureName, Options, Error)
#
#     def SanityCheck(self, SwItemToCheck=defaultNamedNotOptArg,
#                     P1=defaultNamedNotOptArg, P2=defaultNamedNotOptArg):
#         'Sanity check to determine if ok to run addin application'
#         return self._oleobj_.InvokeTypes(96, LCID, 1, (11, 0),
#                                          ((3, 1), (16387, 1), (16387, 1)),
#                                          SwItemToCheck
#                                          , P1, P2)
#
#     def SanityCheck4(self, SwItemToCheck=defaultNamedNotOptArg,
#                      P1=defaultNamedNotOptArg, P2=defaultNamedNotOptArg,
#                      P3=pythoncom.Missing):
#         'Sanity Check'
#         return self._ApplyTypes_(312, 1, (11, 0),
#                                  ((3, 1), (16387, 1), (16387, 1), (16404, 2)),
#                                  'SanityCheck4', None, SwItemToCheck
#                                  , P1, P2, P3)
#
#     def SanityCheck5(self, SwItemToCheck=defaultNamedNotOptArg,
#                      P1=defaultNamedNotOptArg, P2=defaultNamedNotOptArg,
#                      P3=pythoncom.Missing):
#         'Sanity Check'
#         return self._ApplyTypes_(326, 1, (11, 0),
#                                  ((3, 1), (16387, 1), (16387, 1), (16392, 2)),
#                                  'SanityCheck5', None, SwItemToCheck
#                                  , P1, P2, P3)
#
#     def SanityCheck6(self, SwItemToCheck=defaultNamedNotOptArg,
#                      P1=defaultNamedNotOptArg, P2=defaultNamedNotOptArg,
#                      P3=pythoncom.Missing
#                      , P4=pythoncom.Missing, P5=pythoncom.Missing):
#         'Sanity Check'
#         return self._ApplyTypes_(325, 1, (11, 0), (
#             (3, 1), (16387, 1), (16387, 1), (16404, 2), (16387, 2), (16387, 2)),
#                                  'SanityCheck6', None, SwItemToCheck
#                                  , P1, P2, P3, P4, P5
#                                  )
#
#     def SaveSettings(self, FileName=defaultNamedNotOptArg,
#                      SystemOptions=defaultNamedNotOptArg,
#                      ToolbarLayout=defaultNamedNotOptArg,
#                      KeyboardShortcuts=defaultNamedNotOptArg
#                      , MouseGestures=defaultNamedNotOptArg,
#                      MenuCustomization=defaultNamedNotOptArg,
#                      SavedViews=defaultNamedNotOptArg):
#         'Save Settings'
#         return self._oleobj_.InvokeTypes(319, LCID, 1, (3, 0), (
#             (8, 1), (11, 1), (3, 1), (11, 1), (11, 1), (11, 1), (11, 1)),
#                                          FileName
#                                          , SystemOptions, ToolbarLayout,
#                                          KeyboardShortcuts, MouseGestures,
#                                          MenuCustomization
#                                          , SavedViews)
#
#     def SendMsgToUser(self, Message=defaultNamedNotOptArg):
#         'Sends a message to the interactive use in an information box'
#         return self._oleobj_.InvokeTypes(4, LCID, 1, (24, 0), ((8, 1),), Message
#                                          )
#
#     def SendMsgToUser2(self, Message=defaultNamedNotOptArg,
#                        Icon=defaultNamedNotOptArg,
#                        Buttons=defaultNamedNotOptArg):
#         'Send a message to the user'
#         return self._oleobj_.InvokeTypes(76, LCID, 1, (3, 0),
#                                          ((8, 1), (3, 1), (3, 1)), Message
#                                          , Icon, Buttons)
#
#     def SetAddinCallbackInfo(self, ModuleHandle=defaultNamedNotOptArg,
#                              AddinCallbacks=defaultNamedNotOptArg,
#                              Cookie=defaultNamedNotOptArg):
#         'Set addin callback commands.'
#         return self._oleobj_.InvokeTypes(146, LCID, 1, (11, 0),
#                                          ((3, 1), (9, 1), (3, 1)), ModuleHandle
#                                          , AddinCallbacks, Cookie)
#
#     def SetAddinCallbackInfo2(self, ModuleHandle=defaultNamedNotOptArg,
#                               AddinCallbacks=defaultNamedNotOptArg,
#                               Cookie=defaultNamedNotOptArg):
#         'Set addin callback commands.'
#         return self._oleobj_.InvokeTypes(310, LCID, 1, (11, 0),
#                                          ((20, 1), (9, 1), (3, 1)), ModuleHandle
#                                          , AddinCallbacks, Cookie)
#
#     def SetApplySelectionFilter(self, State=defaultNamedNotOptArg):
#         'Set Apply Selection Filter Status'
#         return self._oleobj_.InvokeTypes(110, LCID, 1, (24, 0), ((11, 1),),
#                                          State
#                                          )
#
#     def SetCurrentWorkingDirectory(self,
#                                    CurrentWorkingDirectory=defaultNamedNotOptArg):
#         'Sets the current working directory'
#         return self._oleobj_.InvokeTypes(87, LCID, 1, (11, 0), ((8, 1),),
#                                          CurrentWorkingDirectory
#                                          )
#
#     def SetMissingReferencePathName(self, FileName=defaultNamedNotOptArg):
#         'Set the missing reference file name - used along with ReferenceNotFoundNotify()'
#         return self._oleobj_.InvokeTypes(141, LCID, 1, (24, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def SetMouseDragMode(self, Command=defaultNamedNotOptArg):
#         'Sets the currently running command'
#         return self._oleobj_.InvokeTypes(144, LCID, 1, (11, 0), ((3, 1),),
#                                          Command
#                                          )
#
#     def SetMultipleFilenamesPrompt(self, FileName=defaultNamedNotOptArg):
#         'Set the new file names to open in response to PromptForMultipleFileNamesNotify'
#         return self._oleobj_.InvokeTypes(252, LCID, 1, (24, 0), ((12, 1),),
#                                          FileName
#                                          )
#
#     def SetNewFilename(self, FileName=defaultNamedNotOptArg):
#         'Set the new file to open in response to PromptForFileNewPreNotify'
#         return self._oleobj_.InvokeTypes(264, LCID, 1, (11, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def SetOptions(self, Message=defaultNamedNotOptArg):
#         'Set SOLIDWORKS internal options'
#         return self._oleobj_.InvokeTypes(40, LCID, 1, (11, 0), ((8, 1),),
#                                          Message
#                                          )
#
#     def SetPromptFilename(self, FileName=defaultNamedNotOptArg):
#         'Set the new file to open in response to PromptForFileOpenNotify'
#         return self._oleobj_.InvokeTypes(145, LCID, 1, (24, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def SetPromptFilename2(self, FileName=defaultNamedNotOptArg,
#                            ConfigName=defaultNamedNotOptArg):
#         'Set the new file to open in specific configuration response to PromptForFileOpenNotify'
#         return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0),
#                                          ((8, 1), (8, 1)), FileName
#                                          , ConfigName)
#
#     def SetSearchFolders(self, FolderType=defaultNamedNotOptArg,
#                          Folders=defaultNamedNotOptArg):
#         'Set Search Folders'
#         return self._oleobj_.InvokeTypes(43, LCID, 1, (11, 0), ((3, 1), (8, 1)),
#                                          FolderType
#                                          , Folders)
#
#     def SetSelectionFilter(self, SelType=defaultNamedNotOptArg,
#                            State=defaultNamedNotOptArg):
#         'Set selection filter status'
#         return self._oleobj_.InvokeTypes(90, LCID, 1, (24, 0),
#                                          ((3, 1), (11, 1)), SelType
#                                          , State)
#
#     def SetSelectionFilters(self, SelType=defaultNamedNotOptArg,
#                             State=defaultNamedNotOptArg):
#         'Set status for multiple selection filters'
#         return self._oleobj_.InvokeTypes(108, LCID, 1, (24, 0),
#                                          ((12, 1), (11, 1)), SelType
#                                          , State)
#
#     def SetThirdPartyPopupMenuState(self, RegisterId=defaultNamedNotOptArg,
#                                     IsActive=defaultNamedNotOptArg):
#         'Set third party popup menu state to active or non active'
#         return self._oleobj_.InvokeTypes(286, LCID, 1, (11, 0),
#                                          ((3, 1), (11, 1)), RegisterId
#                                          , IsActive)
#
#     def SetToolbarDock(self, ModuleIn=defaultNamedNotOptArg,
#                        ToolbarIDIn=defaultNamedNotOptArg,
#                        DocStatePosIn=defaultNamedNotOptArg):
#         'Set the docking state of the toolbar'
#         return self._oleobj_.InvokeTypes(131, LCID, 1, (24, 0),
#                                          ((8, 1), (3, 1), (3, 1)), ModuleIn
#                                          , ToolbarIDIn, DocStatePosIn)
#
#     def SetToolbarDock2(self, Cookie=defaultNamedNotOptArg,
#                         ToolbarId=defaultNamedNotOptArg,
#                         DockingState=defaultNamedNotOptArg):
#         'Set the docking state of the toolbar'
#         return self._oleobj_.InvokeTypes(155, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (3, 1)), Cookie
#                                          , ToolbarId, DockingState)
#
#     def SetToolbarVisibility(self, Toolbar=defaultNamedNotOptArg,
#                              Visibility=defaultNamedNotOptArg):
#         'Set toolbar visibility'
#         return self._oleobj_.InvokeTypes(278, LCID, 1, (24, 0),
#                                          ((3, 1), (11, 1)), Toolbar
#                                          , Visibility)
#
#     def SetUserPreferenceDoubleValue(self,
#                                      UserPreferenceValue=defaultNamedNotOptArg,
#                                      Value=defaultNamedNotOptArg):
#         "Set the System's User Preference Double Value"
#         return self._oleobj_.InvokeTypes(47, LCID, 1, (11, 0), ((3, 1), (5, 1)),
#                                          UserPreferenceValue
#                                          , Value)
#
#     def SetUserPreferenceIntegerValue(self,
#                                       UserPreferenceValue=defaultNamedNotOptArg,
#                                       Value=defaultNamedNotOptArg):
#         'Set User Preference Integer Value'
#         return self._oleobj_.InvokeTypes(51, LCID, 1, (11, 0), ((3, 1), (3, 1)),
#                                          UserPreferenceValue
#                                          , Value)
#
#     def SetUserPreferenceStringListValue(self,
#                                          UserPreference=defaultNamedNotOptArg,
#                                          Value=defaultNamedNotOptArg):
#         'Set User Preference String List Value'
#         return self._oleobj_.InvokeTypes(67, LCID, 1, (24, 0), ((3, 1), (8, 1)),
#                                          UserPreference
#                                          , Value)
#
#     def SetUserPreferenceStringValue(self, UserPreference=defaultNamedNotOptArg,
#                                      Value=defaultNamedNotOptArg):
#         'Set User Preference String Value'
#         return self._oleobj_.InvokeTypes(120, LCID, 1, (11, 0),
#                                          ((3, 1), (8, 1)), UserPreference
#                                          , Value)
#
#     def SetUserPreferenceToggle(self, UserPreferenceValue=defaultNamedNotOptArg,
#                                 OnFlag=defaultNamedNotOptArg):
#         'Set User Preference Toggle'
#         return self._oleobj_.InvokeTypes(45, LCID, 1, (24, 0),
#                                          ((3, 1), (11, 1)), UserPreferenceValue
#                                          , OnFlag)
#
#     def ShowBubbleTooltip(self, PointAt=defaultNamedNotOptArg,
#                           FlashButtonIDs=defaultNamedNotOptArg,
#                           TitleResID=defaultNamedNotOptArg,
#                           TitleString=defaultNamedNotOptArg
#                           , MessageString=defaultNamedNotOptArg):
#         'Show bubble tooltip given the SOLIDWORKS resource id'
#         return self._oleobj_.InvokeTypes(192, LCID, 1, (24, 0), (
#             (3, 1), (8, 1), (3, 1), (8, 1), (8, 1)), PointAt
#                                          , FlashButtonIDs, TitleResID,
#                                          TitleString, MessageString)
#
#     def ShowBubbleTooltipAt(self, PointX=defaultNamedNotOptArg,
#                             PointY=defaultNamedNotOptArg,
#                             ArrowPos=defaultNamedNotOptArg,
#                             TitleString=defaultNamedNotOptArg
#                             , MessageString=defaultNamedNotOptArg,
#                             UrlLoc=defaultNamedNotOptArg):
#         'Show a bubble tip at a given screen coordinate.'
#         return self._oleobj_.InvokeTypes(198, LCID, 1, (24, 0), (
#             (3, 1), (3, 1), (3, 1), (8, 1), (8, 1), (8, 1)), PointX
#                                          , PointY, ArrowPos, TitleString,
#                                          MessageString, UrlLoc
#                                          )
#
#     def ShowBubbleTooltipAt2(self, PointX=defaultNamedNotOptArg,
#                              PointY=defaultNamedNotOptArg,
#                              ArrowPos=defaultNamedNotOptArg,
#                              TitleString=defaultNamedNotOptArg
#                              , MessageString=defaultNamedNotOptArg,
#                              TitleBitmapID=defaultNamedNotOptArg,
#                              TitleBitmap=defaultNamedNotOptArg,
#                              UrlLoc=defaultNamedNotOptArg,
#                              Cookie=defaultNamedNotOptArg
#                              , LinkStringID=defaultNamedNotOptArg,
#                              LinkString=defaultNamedNotOptArg,
#                              CallBack=defaultNamedNotOptArg):
#         'Show a bubble tip at a given screen coordinate.'
#         return self._oleobj_.InvokeTypes(289, LCID, 1, (24, 0), (
#             (3, 1), (3, 1), (3, 1), (8, 1), (8, 1), (3, 1), (8, 1), (8, 1),
#             (3, 1),
#             (3, 1), (8, 1), (8, 1)), PointX
#                                          , PointY, ArrowPos, TitleString,
#                                          MessageString, TitleBitmapID
#                                          , TitleBitmap, UrlLoc, Cookie,
#                                          LinkStringID, LinkString
#                                          , CallBack)
#
#     def ShowHelp(self, HelpFile=defaultNamedNotOptArg,
#                  HelpTopic=defaultNamedNotOptArg):
#         'Show a help topic'
#         return self._oleobj_.InvokeTypes(224, LCID, 1, (24, 0),
#                                          ((8, 1), (3, 1)), HelpFile
#                                          , HelpTopic)
#
#     def ShowThirdPartyPopupMenu(self, RegisterId=defaultNamedNotOptArg,
#                                 Posx=defaultNamedNotOptArg,
#                                 Posy=defaultNamedNotOptArg):
#         'Show third party popup menu'
#         return self._oleobj_.InvokeTypes(282, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1), (3, 1)), RegisterId
#                                          , Posx, Posy)
#
#     def ShowToolbar(self, ModuleName=defaultNamedNotOptArg,
#                     ToolbarId=defaultNamedNotOptArg):
#         'Shows a toolbar'
#         return self._oleobj_.InvokeTypes(62, LCID, 1, (11, 0), ((8, 1), (3, 1)),
#                                          ModuleName
#                                          , ToolbarId)
#
#     def ShowToolbar2(self, Cookie=defaultNamedNotOptArg,
#                      ToolbarId=defaultNamedNotOptArg):
#         'Shows a toolbar'
#         return self._oleobj_.InvokeTypes(151, LCID, 1, (11, 0),
#                                          ((3, 1), (3, 1)), Cookie
#                                          , ToolbarId)
#
#     def ShowTooltip(self, ToolbarName=defaultNamedNotOptArg,
#                     ButtonID=defaultNamedNotOptArg,
#                     SelectIDMask1=defaultNamedNotOptArg,
#                     SelectIDMask2=defaultNamedNotOptArg
#                     , TitleString=defaultNamedNotOptArg,
#                     MessageString=defaultNamedNotOptArg):
#         'Show tooltips for selected toolbar buttons'
#         return self._oleobj_.InvokeTypes(212, LCID, 1, (24, 0), (
#             (8, 1), (3, 1), (3, 1), (3, 1), (8, 1), (8, 1)), ToolbarName
#                                          , ButtonID, SelectIDMask1,
#                                          SelectIDMask2, TitleString,
#                                          MessageString
#                                          )
#
#     def SolidWorksExplorer(self):
#         'Starts the SOLIDWORKS Explorer'
#         return self._oleobj_.InvokeTypes(118, LCID, 1, (24, 0), (), )
#
#     def UnInstallQuickTipGuide(self, PInterface=defaultNamedNotOptArg):
#         'UnInstalls a quickTip guide'
#         return self._oleobj_.InvokeTypes(200, LCID, 1, (24, 0), ((9, 1),),
#                                          PInterface
#                                          )
#
#     def UnloadAddIn(self, FileName=defaultNamedNotOptArg):
#         'Unload an Add-In DLL'
#         return self._oleobj_.InvokeTypes(79, LCID, 1, (3, 0), ((8, 1),),
#                                          FileName
#                                          )
#
#     def VersionHistory(self, FileName=defaultNamedNotOptArg):
#         'Get the version history of the given file'
#         return self._ApplyTypes_(81, 1, (12, 0), ((8, 1),), 'VersionHistory',
#                                  None, FileName
#                                  )
#
#     _prop_map_get_ = {
#         "ActiveDoc": (1, 2, (9, 0), (), "ActiveDoc", None),
#         "ActivePrinter": (48, 2, (8, 0), (), "ActivePrinter", None),
#         "ApplicationType": (330, 2, (3, 0), (), "ApplicationType", None),
#         "CommandInProgress": (228, 2, (11, 0), (), "CommandInProgress", None),
#         "EnableBackgroundProcessing": (
#             287, 2, (11, 0), (), "EnableBackgroundProcessing", None),
#         "EnableFileMenu": (265, 2, (11, 0), (), "EnableFileMenu", None),
#         "FrameHeight": (261, 2, (3, 0), (), "FrameHeight", None),
#         "FrameLeft": (258, 2, (3, 0), (), "FrameLeft", None),
#         "FrameState": (262, 2, (3, 0), (), "FrameState", None),
#         "FrameTop": (259, 2, (3, 0), (), "FrameTop", None),
#         "FrameWidth": (260, 2, (3, 0), (), "FrameWidth", None),
#         # Method 'IActiveDoc' returns object of type 'IModelDoc'
#         "IActiveDoc": (16, 2, (9, 0), (), "IActiveDoc",
#                        '{83A33D46-27C5-11CE-BFD4-00400513BB57}'),
#         # Method 'IActiveDoc2' returns object of type 'IModelDoc2'
#         "IActiveDoc2": (156, 2, (9, 0), (), "IActiveDoc2",
#                         '{B90793FB-EF3D-4B80-A5C4-99959CDB6CEB}'),
#         # Method 'JournalManager' returns object of type 'IJournalManager'
#         "JournalManager": (246, 2, (9, 0), (), "JournalManager",
#                            '{338D2790-A47F-45BC-AA03-E70B711CA811}'),
#         "StartupProcessCompleted": (
#             311, 2, (11, 0), (), "StartupProcessCompleted", None),
#         "TaskPaneIsPinned": (232, 2, (11, 0), (), "TaskPaneIsPinned", None),
#         "UserControl": (28, 2, (11, 0), (), "UserControl", None),
#         "UserControlBackground": (
#             210, 2, (11, 0), (), "UserControlBackground", None),
#         "UserTypeLibReferences": (
#             205, 2, (12, 0), (), "UserTypeLibReferences", None),
#         "Visible": (27, 2, (11, 0), (), "Visible", None),
#     }
#     _prop_map_put_ = {
#         "ActivePrinter": ((48, LCID, 4, 0), ()),
#         "CommandInProgress": ((228, LCID, 4, 0), ()),
#         "EnableBackgroundProcessing": ((287, LCID, 4, 0), ()),
#         "EnableFileMenu": ((265, LCID, 4, 0), ()),
#         "FrameHeight": ((261, LCID, 4, 0), ()),
#         "FrameLeft": ((258, LCID, 4, 0), ()),
#         "FrameState": ((262, LCID, 4, 0), ()),
#         "FrameTop": ((259, LCID, 4, 0), ()),
#         "FrameWidth": ((260, LCID, 4, 0), ()),
#         "TaskPaneIsPinned": ((232, LCID, 4, 0), ()),
#         "UserControl": ((28, LCID, 4, 0), ()),
#         "UserControlBackground": ((210, LCID, 4, 0), ()),
#         "UserTypeLibReferences": ((205, LCID, 4, 0), ()),
#         "Visible": ((27, LCID, 4, 0), ()),
#     }
#
#     def __iter__(self):
#         "Return a Python iterator for this object"
#         try:
#             ob = self._oleobj_.InvokeTypes(-4, LCID, 3, (13, 10), ())
#         except pythoncom.error:
#             raise TypeError("This object does not support enumeration")
#         return win32com.client.util.Iterator(ob, None)
