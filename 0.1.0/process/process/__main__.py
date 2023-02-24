import os
import inspect
import traceback
from time import sleep
from pprint import pprint


class TypeMismatchDataIn2DataOut(Exception):
    pass


class ProcessFailed2Initalize(Exception):
    pass


class Failed2ConvertParameterValueType(Exception):
    pass


class Failed2ParseParameterValue(Exception):
    pass


class TooManyFunctionParameters(Exception):
    pass


class Console:
    """_The console interface to the Process Class._

    Args:
        Process.Console (_Process.shared.Console_): _Process.shared.Console_
    """

    # [TODO] For menu ordering. Does it depend on the
    # order in which functions appear in this file,
    # and in the inheriting Process file?
    # Should functions get listed by inspect as found?
    __all__ = ['process_menu', 'developer_help'
               'shared_documents','create_table'
               'shutdown','create_new_table'
               'sharepoint_xlsx_to_table',
               'local_file_to_table', 
               "setup_text_msg_alerts", 
               "submit_a_bug_report", 
               "submit_a_feature_request", 
               "go_back_to_main", "shutdown"]

    GET_OBJMEM_NAME = lambda m: m[0]
    GET_OBJMEM_TYPE = lambda m: str(m[1]).split()[0].strip()[1:]
    
    def __init__(self, Process, **kwargs):
        self._dataIn = None # placeholder for return from .dataIn
        
        if kwargs:
            for kwarg in kwargs:
                setattr(self, kwarg, kwargs.get(kwarg))
        
        self.api = Process()
        try:
            self.api.console.center_to_screen("DXC")
            self.api.console.center_to_screen("O365 Desktop Automator")
            # [NOTE] There can be no direct imports from the main package
            # into this file. The version number here must be updated
            # manually.
            self.api.console.center_to_screen("v0.1.0")

        except:
            raise ProcessFailed2Initalize

        opt = 0 # option number for Process Menu
        self.OPTIONS["Process Menu"]= dict()
        
        for member in inspect.getmembers(self.api):
            name = self.GET_OBJMEM_NAME(member)
        
            if not name[0] == "_":
                func = self.GET_OBJMEM_TYPE(member)
        
                if func == 'function' and not name in self.__all__:
                    opt += 1 # starting with 1 provides a better ux.
                    functionAsTitledName = "".join([w.title() for w in name.split("_")])
                    self.OPTIONS["Process Menu"][opt] = functionAsTitledName
        
        opts = len(self.OPTIONS["Process Menu"])
        self.OPTIONS["Process Menu"][opts+1] = "Go Back To Main"

    def __process_fail(self, exception: Exception):
        # Does not return to a menu as it is unknown from which menu this process
        # was called, the calling method should redirect to the previous menu after 
        # a call to .__process_fail has been made.
        # Clears the screen of any error messages first. option for traceback readout.
        # The process function could fail for a number of reasons before and after
        # it has been placed into production. For example, only one instance of 
        # Excel can be opened at a time (without Administrator priviledges), As 
        # vba script execution is running via a separate process, the user could
        # inadvertantly try to execute multiple scripts, of multiple instances of
        # the same script at the same time. In such instances, the functions related
        # to the application running the destkop script would show the message
        # below when selected. Other functions, streamed dataframe operations
        # from SharePoint, would not be impacted.
        self.api.console.clear()
        pprint("This process is currently unavaiable.")
        
        if self.api.console.get_yesno("Would you like to read the traceback?") == "yes":    
            try:
                self.api.console.clear()
                traceback.print_exception(exception)
                # without a pause here, the screen would be cleared 
                # by options selection.
                _ = input("Continue? [<Enter>]")
            
            except:
                print(exception)
            
            if self.api.console.get_yesno("Continue?") == "no":
                os._exit(0)

    def __process_menu(self, selection: str):
        # Parameters must be fetched from the command line using the tools
        # provided through self.console if the process is being run from 
        # the console, with the exception of .dataOut (which should only include
        # parameters are are equal to the return from .dataIn).

        # Args: 
        #   selection, _str_ : _the selection option value_ 
        titledNameAsFunction = '_'.join([w.lower() for w in self.OPTIONS["Process Menu"][selection]])
        if titledNameAsFunction == "Process Menu":
            return self.process_menu()
        elif titledNameAsFunction == "Console Menu":
            return self.console_menu()
        
        func = eval(f"self.api.{titledNameAsFunction}")
        try:
            func((self._dataIn))
        
        except:
            parameters_types = func.__annotations__
            if len(parameters_types) != len(self._dataIn):
                self.__process_fail(TooManyFunctionParameters)

            parameters = list() # short for parameters
            for name, t in parameters_types.items():
                value = input(f"{name}?: ").strip()
                
                if t == dict:
                    param = dict()
                    value.replace("{","")
                    value.replace("}","")
                    pairs = [kv.strip() for kv in value.split(",")]
                    pairs = [kv.split(":") for kv in pairs]
                
                    try:
                        for pair in pairs:
                            k, v = pair[0].strip(), pair[1].strip()
                            param[k] = v
                        parameters.append(tuple(name,param))
                
                    except:       
                        self.__process_fail(Failed2ParseParameterValue)

                elif t in [list, tuple]:
                    value.replace("[","")
                    value.replace("]","")
                    value.replace("(","")
                    value.replace(")","")

                    value = [v.strip() for v in value.split(",")]
                    parameters[name] = value if t == list else tuple(value)

                elif t in [str, int, float]:
                    try:
                        parameters.append(name, t(value))
                    
                    except:
                        self.__process_fail(Failed2ConvertParameterValueType)
                
                else:
                    self.__process_fail(Failed2ConvertParameterValueType)

            func((parameters.values))
        
        return self.process_menu()

    def process_menu(self):
        menu = self.OPTIONS["Process Menu"]
        sel = self.console.get_option_selection(menu)
        
        selection = self.OPTIONS["Process Menu"][sel]
        if selection == "Back To Main": return self.main()
        return self.__process_menu(selection)

    def create_table(self):
        menu = {
            1: "Console To Table",
            2: "Local File To Table",
            3: "SharePoint File To Table",
            4: "Encrypted SharePoint Xlsx To Table",
            5: "Go Back to Console Menu"}
        sel = self.console.get_option_selection(menu)
        if sel != 5:
            df = self.__exec_console_submenu_function(menu[sel])
            self.__table_complete(df)

    def setup_text_msg_alerts(self):
        while True:
            try:
                areacode = int(input("Area Code [###]: ").strip())
                phonenum = input("Phone Number [###-####]: ")
                phonenum = "".join([i for i in list(phonenum) if i.isnumeric()])
                phonenum = areacode + phonenum
                if not len(phonenum) == 10:
                    raise ValueError
                break
            
            except:
                print("???\033[0K\r")
                sleep(2)

        options = {
            1 : "AT&T",
            2 : "T-Mobile",
            3 : "Verizon",
            4 : "Other"}  
        sel = self.api.console.get_option_selection(options)
        
        msg = "Other carriers are not yet available. "
        msg += "Please submit a feature request from the startup menu with your carrier name."
        if sel == 4: 
            pprint(msg)
            return

        textemail = {
            1: f"{phonenum}@rtxt.att.net",
            2: f"{phonenum}@tmomoail.net",
            3: f"{phonenum}@vtext.com"}

        # [TODO] Store text alert address
        textAlertAddress = textemail[sel]

    def send_text_alert(self, msg: str):
        # [TODO] add config file load into kwargs from main__.__main__
        # might need to encrypt if credentials are stored...
        if hasattr(self, "textAlertAddress"):
            self.api.outlook.send_email(body=msg,to=self.textAlertAddress)
        
        else:
            pprint("Please 'Setup Text Msg Alerts' first from the startup menu.")

    def exec_console_function(self, selection: str):
        titledNameAsFunction = '_'.join([w.lower() for w in selection])
        return eval(f"self.{titledNameAsFunction}")()

    def exec_console_submenu_function(self, selection: str):
        titledNameAsFunction = '_'.join([w.lower() for w in selection])
        return eval(f"self.__{titledNameAsFunction}")()
   
    def console_menu(self):
        """_summary_

        Returns:
            _type_: _description_
        """

        name = "Console Menu"
        menu = self.options[name]
        sel = self.console.get_option_selection(menu)
        
        selection = self.options[name][sel]
        self.__exec_console_function(name, self.options[name][selection])
        return self.console_menu()

    def main_menu(self):
        """_summary_

        Returns:
            _type_: _description_
        """

        name = "Main Menu"
        menu = self.options[name]
        sel = self.console.get_option_selection(menu)

        selection = self.options[name][sel]
        self.__exec_function(name, self.options[name][selection])
        return self.main_menu()
    
    def developer_help(self):
        pass
    
    def developer_docs(self):
        pass

    def shutdown(self):
        os._exit(0)


if __name__ == "__main__":

    # [NOTE] This is the only way the inheriting Console class
    #  can have no specifiable content, although it could if the
    #  user wanted more options displayed on the main screen. This too should
    #  be left as is.
    from .__init__ import Process 

    try:
        app = Console(Process)
        app.main()
    
    except Exception as e:
        traceback.print_exception(e)