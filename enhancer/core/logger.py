"""
Logging, script path resolution, and exception formatting utilities.
"""

import os
import sys
import logging
import traceback
from datetime import datetime


def script_path():
    """
    get the script path of the file that starts the program
    """
    current_script = os.path.realpath(sys.argv[0])  # script directory
    return os.path.splitext(current_script)[0]  # name, ext


class OperationDebugger(logging.Logger):
    _is_setting_up = False

    def __init__(self, file_name=None, level=logging.DEBUG, stream=None, formatter=None):
        """
        create a file debugger to log relevant info, errors and debug data to a log.

        :param file_name: name of debugger and file
        :param level: logging level. logging.DEBUG by default
        :param stream: use a stream to print to. None by default
        :return:
        """
        self._is_setting_up = True
        self._file_created = False

        file_timestamp = '%Y-%m-%d %H:%M:%S'
        if formatter is None:
            formatter = logging.Formatter(
                fmt='%(asctime)s# %(levelname)-8s# %(message)s',
                datefmt=file_timestamp
            )
        now = datetime.now()
        if file_name is None:
            file_name = '{name}-{timestamp}.log'.format(
                timestamp=now.strftime(file_timestamp).replace(":", "_"),
                name=script_path()
            )
        self.formatter = formatter

        super().__init__(file_name, level)

        if stream is not None:
            screen_handler = logging.StreamHandler(stream=stream)
            screen_handler.setFormatter(formatter)
            self.addHandler(screen_handler)

        self._is_setting_up = False

    def __getattribute__(self, item):
        # register this only when logging starts
        if not object.__getattribute__(self, "_is_setting_up"):
            if not object.__getattribute__(self, "_file_created"):  # this must run once
                self._is_setting_up = True
                file_name = object.__getattribute__(self, "name")
                formatter = object.__getattribute__(self, "formatter")
                add_handler_fn = object.__getattribute__(self, "addHandler")
                handler = logging.FileHandler(file_name, mode='w')
                handler.setFormatter(formatter)
                add_handler_fn(handler)
                self._file_created = True
                self._is_setting_up = False
        return super().__getattribute__(item)


file_debugger = OperationDebugger()


def exception_to_stream(e, stream_handler=None, template=None):
    """
    prints Exception message to stream much like traceback.format_exception or traceback.print_exception

    :param e: Exception
    :param stream_handler:
    :param template:
    :return:
    """
    # gotten from book "The Python 3 Standard Library by Example: Pyth 3 Stan Libr Exam _2"
    template = template or "File \"{filename:<23}\", line {linenum}, in {funcname}\n{source}"
    lines = []
    for tb_info in traceback.extract_tb(e.__traceback__):
        filename, linenum, funcname, source = tb_info
        filename = os.path.basename(filename)
        if funcname != '<module>':
            funcname += '()'
        msg = template.format(filename=filename, linenum=linenum, funcname=funcname, source=source)
        lines.append(msg)

    if stream_handler:
        stream_handler("\n".join(lines))
    return lines

