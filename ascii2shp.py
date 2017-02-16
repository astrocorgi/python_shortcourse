"""ascii2shp.py - python wrapper around ascii2shp.exe that allows it to be called as an ArcGIS geoprocessing script

Copyright (C) 2006, Jason J. Roberts

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software (the "Software"), to deal in the Software without
restriction, including without limitation the rights to use, copy, modify,
merge, publish, distribute, sublicense, and/or sell copies of the
Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.
"""

# This script borrows code from Marine Ecology Tools, one of my other Python
# projects. That project also includes an ascii2shp wrapper.


import os
import sys
import traceback
import win32com.client


class Geoprocessor:
    """Class that allocates a singleton ArcGIS geoprocessor object and returns it to callers. Probably not thread safe. Doesn't fail gracefully."""

    instance = win32com.client.Dispatch("esriGeoprocessing.GPDispatch.1")


class MarineError(Exception):
    """Exception raised by all Marine Ecology Tools."""

    def __init__(self, message, indent_level):
        self.message = message
        try:
            indent = " " * indent_level
            message = indent + message.replace("\n", "\n" + indent)
            Geoprocessor.instance.AddError(message)
        except:
            pass

    def __repr__(self):
        return self.message

    def __str__(self):
        return repr(self)


class MarineTool:
    """Base class for all Marine Ecology Tools."""

    def __init__(self, product=None, extensions=None, keyword_args=None):
        """Base class initialization code. Should be called at the top of the derived class's init method."""

        self._product = product
        self._extensions = extensions
        self._keyword_args = keyword_args
        self._command_line_arg_count = 0

        if keyword_args is not None and keyword_args.has_key("indent_level"):
            self._indent_level = keyword_args["indent_level"]
        else:
            self._indent_level = 1

        if keyword_args is not None and keyword_args.has_key("from_command_line"):
            self._from_command_line = keyword_args["from_command_line"]
        else:
            self._from_command_line = 0

        if keyword_args is not None and keyword_args.has_key("verbose_logging"):
            self._verbose_logging = keyword_args["verbose_logging"]
        else:
            self._verbose_logging = 0

        if self._extensions is not None and not isinstance(self._extensions, types.StringTypes) and not isinstance(self._extensions, types.ListType):
            raise MarineError("The extensions parameter must be a string specifying exactly one extension or a list of strings that each specify exactly one extension.", self._indent_level)
        
        self.tool_name = str(self.__class__)
        if (self.tool_name.rfind(".") >= 0):
            self.tool_name = self.tool_name[self.tool_name.rfind(".")+1:]

        self._log_verbose(self.tool_name + " started.")
        self._indent_level= self._indent_level + 1
        self._log_verbose("Python " + sys.version)

    def run(self):
        """Initializes and runs this tool."""

        try:
            checkin = 0
            if self._product is not None:
                self._set_product()
            if self._extensions is not None:
                checkin = 1
                self._checkout_extensions()

            result = None
            try:
                result = self._run_tool()
            finally:
                if checkin:
                    self._checkin_extensions()

            self._indent_level = self._indent_level - 1
            self._log_verbose(self.tool_name + " completed successfully.")
            return result
        except:
            self._indent_level = self._indent_level - 1
            self._log(self.tool_name + " failed.", "error")
            raise

    def _run_tool(self):
        """Derived classes must implement this method."""
        
        raise MarineError("The " + self.tool_name + " tool has not been implemented.", self._indent_level)

    def _log(self, message, level="info"):
        """Helper function for logging."""
        
        try:
            indent = " " * self._indent_level
            message = indent + message.replace("\n", "\n" + indent)
            print(message)
            if level.lower() == "warning":
                Geoprocessor.instance.AddWarning(message)
            elif level.lower() == "error":
                Geoprocessor.instance.AddError(message)
            elif level.lower() == "info" or self._verbose_logging:
                Geoprocessor.instance.AddMessage(message)
        except:
            pass

    def _log_verbose(self, message):
        self._log(message, "verbose")

    def _log_returned_geoprocessor_messages(self, force_verbose_logging=0):
        """Logs messages returned by ArcGIS tools invoked through Geoprocessor.instance (e.g., Geoprocessor.instance.Clip_analysis). Since non-tool GpDispatch methods (e.g. Geoprocessor.instance.UpdateCursor) do not return messages, this will log nothing if invoked after calling them."""\

        self._indent_level = self._indent_level + 1
        
        try:
            i = 0
            while i < Geoprocessor.instance.MessageCount:
                sev = Geoprocessor.instance.GetSeverity(i)
                if sev == 0:
                    if force_verbose_logging:
                        self._log(Geoprocessor.instance.GetMessage(i), "info")
                    else:
                        self._log_verbose(Geoprocessor.instance.GetMessage(i))
                elif sev == 1:
                    self._log(Geoprocessor.instance.GetMessage(i), "warning")
                elif sev == 2:
                    self._log(Geoprocessor.instance.GetMessage(i), "error")
                i = i + 1
        except:
            pass
        
        self._indent_level = self._indent_level - 1

    def _parse_string_arg(self, arg_name, default=None, required=0):
        value = None
        if self._from_command_line:
            self._command_line_arg_count = self._command_line_arg_count + 1
            value = sys.argv[self._command_line_arg_count]
        elif self._keyword_args is not None and self._keyword_args.has_key(arg_name):
            value = self._keyword_args[arg_name]
            if value is not None and not isinstance(value, types.StringTypes):
                raise MarineError("Parameter " + arg_name + " must be a string.", self._indent_level)
        if value is not None:
            value = value.strip()
            if value == "#":
                value = None
        if value is None:
            if required:
                if self._from_command_line:
                    raise MarineError("Parameter " + str(self._command_line_arg_count) + ", " + arg_name + ", may not be omitted. Please provide a value for this parameter.", self._indent_level)
                else:
                    raise MarineError("Parameter " + arg_name + " may not be omitted. Please provide a value for this parameter.", self._indent_level)
            if default is not None:
                if not isinstance(default, types.StringTypes):
                    raise MarineError("Internal error in the " + self.tool_name + " tool. The \"default\" parameter to self._parse_string_arg() must be a string. Please contact the developer of the tool.", self._indent_level)
                value = default
        self.__dict__[arg_name] = value
        self._log_verbose(arg_name + " = " + str(value))

    def _parse_boolean_arg(self, arg_name, default=0, required=0):
        value = None
        if self._from_command_line:
            self._command_line_arg_count = self._command_line_arg_count + 1
            arg = sys.argv[self._command_line_arg_count]
            if arg is None or len(arg.strip()) <= 0 or arg.strip() == "#":
                if required:
                    raise MarineError("Parameter " + str(self._command_line_arg_count) + ", " + arg_name + ", may not be omitted. Please specify either True or False for this parameter.", self._indent_level)
            else:
                if arg.strip().lower() == "true":
                    value = 1
                elif arg.strip().lower() == "false":
                    value = 0
                elif required:
                    raise MarineError("Parameter " + str(self._command_line_arg_count) + ", " + arg_name + ", must be either True or False.", self._indent_level)
                else:
                    raise MarineError("Parameter " + str(self._command_line_arg_count) + ", " + arg_name + ", must be either True, False or #. (# means use the default value, " + str(default) + " for this script. When ArcGIS invokes this script from an ArcGIS geoprocessing model, it provides # automatically when you leave the field blank.)", self._indent_level)
        else:
            if self._keyword_args is not None and self._keyword_args.has_key(arg_name):
                value = self._keyword_args[arg_name]
            if value is not None and not isinstance(value, types.BooleanType):
                raise MarineError("Parameter " + arg_name + " must be a boolean.", self._indent_level)
            if value is None and required:
                raise MarineError("Parameter " + arg_name + " may not be omitted. Please provide a boolean for this parameter.", self._indent_level)
        if value is None and default is not None:
            if not isinstance(default, types.BooleanType):
                raise MarineError("Internal error in the " + self.tool_name + " tool. The \"default\" parameter to self._parse_boolean_arg() must be a boolean. Please contact the developer of the tool.", self._indent_level)
            value = default
        self.__dict__[arg_name] = value
        self._log_verbose(arg_name + " = " + str(value))

    def _parse_int_arg(self, arg_name, default=0, required=0):
        value = None
        if self._from_command_line:
            self._command_line_arg_count = self._command_line_arg_count + 1
            arg = sys.argv[self._command_line_arg_count]
            if arg is None or len(arg.strip()) <= 0 or arg.strip() == "#":
                if required:
                    raise MarineError("Parameter " + str(self._command_line_arg_count) + ", " + arg_name + ", may not be omitted. Please specify an integer ranging from -2147483648 to 2147483647.", self._indent_level)
            else:
                try:
                    value = int(arg.strip())
                except:
                    raise MarineError("\"" + arg.strip() + "\" is an invalid value for parameter " + str(self._command_line_arg_count) + ", " + arg_name + ". Please specify an integer ranging from -2147483648 to 2147483647.", self._indent_level)
        else:
            if self._keyword_args is not None and self._keyword_args.has_key(arg_name):
                value = self._keyword_args[arg_name]
            if value is not None and not isinstance(value, types.IntType):
                raise MarineError("Parameter " + arg_name + " must be an int.", self._indent_level)
            if value is None and required:
                raise MarineError("Parameter " + arg_name + " may not be omitted. Please provide an int for this parameter.", self._indent_level)
        if value is None and default is not None:
            if not isinstance(default, types.IntType):
                raise MarineError("Internal error in the " + self.tool_name + " tool. The \"default\" parameter to self._parse_int_arg() must be an int. Please contact the developer of the tool.", self._indent_level)
            value = default
        self.__dict__[arg_name] = value
        self._log_verbose(arg_name + " = " + str(value))

    def _parse_float_arg(self, arg_name, default=0.0, required=0):
        value = None
        if self._from_command_line:
            self._command_line_arg_count = self._command_line_arg_count + 1
            arg = sys.argv[self._command_line_arg_count]
            if arg is None or len(arg.strip()) <= 0 or arg.strip() == "#":
                if required:
                    raise MarineError("Parameter " + str(self._command_line_arg_count) + ", " + arg_name + ", may not be omitted. Please specify a floating-point number.", self._indent_level)
            else:
                try:
                    value = float(arg.strip())
                except:
                    raise MarineError("\"" + arg.strip() + "\" is an invalid value for parameter " + str(self._command_line_arg_count) + ", " + arg_name + ". Please specify floating point number.", self._indent_level)
        else:
            if self._keyword_args is not None and self._keyword_args.has_key(arg_name):
                value = self._keyword_args[arg_name]
            if value is not None and not isinstance(value, types.FloatType):
                raise MarineError("Parameter " + arg_name + " must be a float.", self._indent_level)
            if value is None and required:
                raise MarineError("Parameter " + arg_name + " may not be omitted. Please provide a float for this parameter.", self._indent_level)
        if value is None and default is not None:
            if not isinstance(default, types.FloatType):
                raise MarineError("Internal error in the " + self.tool_name + " tool. The \"default\" parameter to self._parse_int_arg() must be a float. Please contact the developer of the tool.", self._indent_level)
            value = default
        self.__dict__[arg_name] = value
        if value is not None:
            self._log_verbose("%s = %g" % (arg_name, value))
        else:
            self._log_verbose(arg_name + " = " + str(value))
    
    def _set_product(self):
        """Set the ArcGIS product level."""

        if self._product is not None:
            try:
                status = Geoprocessor.instance.CheckProduct(self._product)
            except Exception, e:
                self._log("Unable to determine if a license is available for the ArcGIS \"" + self._product + "\" product due to a failure in the ArcGIS GpDispatch.CheckProduct function. Please check your ArcGIS licensing configuration. Error details: " + str(e), "error")
                raise

            if status.lower() != "available":
                raise MarineError("The ArcGIS GpDispatch.CheckProduct function reports that an \"" + self._product + "\" product license is not available on this machine (it returned status code \"" + status + "\"). Please check your ArcGIS licensing configuration.", self._indent_level)

            try:
                Geoprocessor.instance.SetProduct(self._product)
                self._log_verbose("Set the ArcGIS product level to \"" + self._product + "\".")
            except Exception, e:
                self._log("Unable to obtain a license for the ArcGIS \"" + self._product + "\" product due to a failure in the ArcGIS GpDispatch.SetProduct function. Please check your ArcGIS licensing configuration. Error details: " + str(e), "error")
                raise

    def _checkout_extensions(self):
        """Check out the ArcGIS extensions."""
        
        if isinstance(self._extensions, types.StringTypes):
            extension_list = [self._extensions]
        else:
            extension_list = self._extensions
        
        try:
            got_extensions = []
            
            for e in extension_list:
                try:
                    status = Geoprocessor.instance.CheckExtension(e)
                except Exception, e:
                    self._log("Unable to determine if a license is available for the ArcGIS \"" + e + "\" extension due to a failure in the ArcGIS GpDispatch.CheckExtension function. Please check your ArcGIS licensing configuration. Error details: " + str(e), "error")
                    raise

                if status.lower() != "available":
                    raise MarineError("The ArcGIS GpDispatch.CheckExtension function reports that a license for the \"" + e + "\" extension is not available on this machine (it returned status code \"" + status + "\"). Please check your ArcGIS licensing configuration.", self._indent_level)

                try:
                    Geoprocessor.instance.CheckOutExtension(e)
                except Exception, e:
                    self._log("Unable to obtain a license for the ArcGIS \"" + e + "\" extension due to a failure in the ArcGIS GpDispatch.CheckOutExtension function. Please check your ArcGIS licensing configuration. Error details: " + str(e), "error")
                    raise

                got_extensions = got_extensions.append(e)
                self._log_verbose("Checked out a license for the ArcGIS \"" + e + "\" extension.")

        except:
            try:
                for e in got_extensions:
                    Geoprocessor.instance.CheckInExtension(e);
            except:
                pass

    def _checkin_extensions(self):
        """Check in the ArcGIS extensions."""
        
        if isinstance(self._extensions, types.StringTypes):
            extension_list = [self._extensions]
        else:
            extension_list = self._extensions
        
        for e in extension_list:
            try:
                Geoprocessor.instance.CheckInExtension(e)
                self._log_verbose("Checked in the license for the ArcGIS \"" + e + "\" extension.")
            except:
                pass

    # Wrapper functions for the ArcGIS GpDispatch object

    def _GpDispatch_Exists(self, obj):
        """Wrapper around GpDispatch.Exists: Tests the existance of the data object."""

        try:
            return Geoprocessor.instance.Exists(obj)
        except Exception, e:
            self._log("Failed to test the existence of \"" + str(obj) + "\". The ArcGIS geoprocessor Exists function raised an error: " + str(e), "error")
            raise


class Ascii2Shp(MarineTool):
    """Real implementation of the ascii2shp.exe wrapper."""

    def __init__(self, **keyword_args):
        MarineTool.__init__(self, keyword_args=keyword_args)

        self._parse_string_arg("input_textfile", required=1)
        self._parse_string_arg("output_shapefile", required=1)
        self._parse_string_arg("xcol", required=1)
        self._parse_string_arg("ycol", required=1)
        self._parse_string_arg("zcol")
        self._parse_string_arg("mcol")
        self._parse_string_arg("delimiter", default="comma")
        self._parse_string_arg("delimiter_character")
        self._parse_string_arg("comment_string")
        self._parse_string_arg("quote_character")
        self._parse_string_arg("locale")
        self._parse_string_arg("coordinate_system")

    def _run_tool(self):
        """Entry point for this tool."""

        # Validate input parameters.

        if not os.path.isfile(self.input_textfile):
            raise MarineError("The input file \"" + str(self.input_textfile) + "\" does not exist or is not accessible. Please check the file path and try again.", self._indent_level)

        if self._GpDispatch_Exists(self.output_shapefile):
            raise MarineError("The shapefile or other GIS object \"" + str(self.output_shapefile) + "\" already exists. Please delete it and try again.", self._indent_level)

        if self.delimiter is not None:
            if self.delimiter.lower() != "comma" and self.delimiter.lower() != "space or tab"  and self.delimiter.lower() != "user specified":
                raise MarineError("The value \"" + self.delimiter + "\" is not valid for the delimiter parameter. It must be either \"comma\", \"space or tab\" or \"user specified\".", self._indent_level)
            self.delimiter = self.delimiter.lower();
            if self.delimiter.lower() == "user specified":
                if self.delimiter_character is None or len(self.delimiter_character) != 1:
                    raise MarineError("The user specified delimiter must be a single character.", self._indent_level)
        
        if self.comment_string is not None and self.comment_string == "\#":
            self.comment_string = "#"

        if self.quote_character is not None:
            if self.quote_character.lower() != "single quote" and self.quote_character.lower() != "double quote":
                raise MarineError("The value \"" + self.quote_character + "\" is not valid for the quote character parameter. It must be either \"single quote\" or \"double quote\".", self._indent_level)
            self.quote_character = self.quote_character.lower();

        # Build the ASCII2SHP.EXE command line to execute.

        cmd = "cmd.exe /c " + os.path.dirname(sys.argv[0]) + "\\ascii2shp.exe \"%s\" \"%s\" %s %s" % (self.input_textfile, self.output_shapefile, self.xcol, self.ycol)

        if self.zcol is not None:
            cmd = cmd + " -z " + self.zcol

        if self.mcol is not None:
            cmd = cmd + " -m " + self.mcol

        if self.delimiter is not None:
            if self.delimiter == "comma":
                cmd = cmd + " -c"
            elif self.delimiter == "space or tab":
                cmd = cmd + " -t"
            elif self.delimiter == "user specified":
                cmd = cmd + " -d " + self.delimiter_character
            else:
                raise MarineError("Programming error in ascii2shp.py. The value \"" + self.delimiter + "\" is an unknown delimiter value. Please contact the author of this tool.", self._indent_level)

        if self.quote_character is not None:
            if self.quote_character == "single quote":
                cmd = cmd + " -a"
            elif self.quote_character == "double quote":
                cmd = cmd + " -q"
            else:
                raise MarineError("Programming error in ascii2shp.py. The value \"" + self.quote_character + "\" is an unknown delimiter value. Please contact the author of this tool.", self._indent_level)

        if self.locale is not None:
            cmd = cmd + " -l " + self.locale

        # Execute ASCII2SHP.EXE

        self._log("Executing '" + cmd + "'...")

        try:
            f = os.popen(cmd)
        except Exception, e:
            self._log("Failed to execute '" + cmd + "'. The python os.popen() function raised the following exception when passed that command: " + str(e), "error")
            raise

        self._indent_level = self._indent_level + 2
        self._log_verbose("Please wait. No progress will be reported until ascii2shp.exe exits...")

        try:
            ascii2shp_output = f.read()
        except Exception, e:
            self._log("Failed to read from the stdout pipe of the ASCII2SHP.EXE child process. The python file.read() function raised the following exception: " + str(e), "error")
            self._indent_level = self._indent_level - 2
            raise

        try:
            ascii2shp_result = f.close()
        except Exception, e:
            self._log("Failed to close the stdout pipe of the ASCII2SHP.EXE child process. The python file.close() function raised the following exception: " + str(e), "error")
            self._indent_level = self._indent_level - 2
            raise

        if ascii2shp_result is not None:
            self._log("")        
            self._log(ascii2shp_output, "error")
            self._indent_level = self._indent_level - 2
            raise MarineError("Ascii2shp.exe failed.", self._indent_level)

        self._log_verbose("")        
        self._log_verbose(ascii2shp_output)        
        self._indent_level = self._indent_level - 2
        self._log("Ascii2shp.exe completed successfully.")

        # Refresh the catalog. This must be done for Arc to detect that the
        # shapefile was created. If it is not done, the DefineProjection
        # function will fail, saying that the shapefile does not exist. It is
        # not clear what should be passed to RefreshCatalog. When I passed
        # self.output_shapefile, it worked fine on my machine but on someone
        # else's machine, it would then fail on the DefineProjection. When I
        # passed os.path.dirname(self.output_shapefile), it would crash python
        # on my machine (it went straight into "we're sorry" prompt for
        # python.exe). The following hack will hopefully not crash python.exe
        # and successfully refresh the catalog on all machines.

        self._log_verbose("Calling the ArcGIS RefreshCatalog function on \"" + self.output_shapefile + "\"...")
        try:
            Geoprocessor.instance.RefreshCatalog(self.output_shapefile)
        except Exception, e:
            self._log("The ArcGIS RefreshCatalog function failed on \"" + self.output_shapefile + "\". Further geoprocessing involving \"" + self.output_shapefile + "\" may fail until you manually refresh the catalog. The exception raised by RefreshCatalog was: " + str(e), "warning")

        self._log_verbose("Calling the ArcGIS RefreshCatalog function on \"" + os.path.os.path.dirname(os.path.dirname(self.output_shapefile)) + "\"...")
        try:
            Geoprocessor.instance.RefreshCatalog(os.path.dirname(os.path.dirname(self.output_shapefile)))
        except Exception, e:
            self._log("The ArcGIS RefreshCatalog function failed on \"" + os.path.dirname(os.path.dirname(self.output_shapefile)) + "\". Further geoprocessing involving \"" + self.output_shapefile + "\" may fail until you manually refresh the catalog. The exception raised by RefreshCatalog was: " + str(e), "warning")

        # Define the projection.

        if self.coordinate_system is not None:
            self._log("Defining the projection of \"" + self.output_shapefile + "\"...")
            
            try:
                Geoprocessor.instance.DefineProjection_management(self.output_shapefile, self.coordinate_system)
            except:
                self._log_returned_geoprocessor_messages(1)
                raise

            self._log_returned_geoprocessor_messages()


def main():
    """Entry point for this tool when invoked from the command line."""

    verbose_logging=(sys.argv[13] is not None and sys.argv[13].lower() == "true")
    
    try:
        t = Ascii2Shp(from_command_line=1, verbose_logging=verbose_logging)
        return t.run()
    except:
        if verbose_logging:
            msg_list = traceback.format_exception(sys.exc_type, sys.exc_value, sys.exc_traceback)
            msg = ""
            for m in msg_list:
                msg = msg + m
            Geoprocessor.instance.AddError(msg)
        raise

    logger.shutdown()

if __name__ == '__main__':
    main()
