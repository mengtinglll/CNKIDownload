"""
This module provides the logging function.
"""

import logging
import logging.handlers
import os
import sys

class Logger(object):
    """The logger factory class. It is a template to help quickly create a log utility.
    Attributes:
    set_conf(log_file, use_stdout, log_level): this is a static method that returns a configured logger.
    get_logger(tag): this is a static method that returns a configured logger.
    """
    __loggers = {}

    __use_stdout = True
    __log_file = ""
    __log_level = logging.DEBUG

    @staticmethod
    def config(log_file, use_stdout, log_level):
        """set the config, where config is a ConfigParser object
        """
        key_logger_section = "logger"
        Logger.__use_stdout = use_stdout
        Logger.__log_level = log_level
        dirname = os.path.dirname(log_file)
        if (not os.path.isfile(log_file)) and (not os.path.isdir(dirname)):
            try:
                os.makedirs(dirname)
            except OSError as e:
                print("create path '%s' for logging failed: %s" % (dirname, e))
                sys.exit()
        Logger.__log_file = log_file

    @staticmethod
    def get_logger(tag):
        """return the configured logger object
        """
        if tag not in Logger.__loggers:
            Logger.__loggers[tag] = logging.getLogger(tag)
            Logger.__loggers[tag].setLevel(Logger.__log_level)
            formatter = logging.Formatter(
                "[%(name)s][%(levelname)s] %(asctime)s "
                "%(filename)s:%(lineno)s %(message)s")
            file_handler = logging.handlers.TimedRotatingFileHandler(
                Logger.__log_file, when='D', interval=1, backupCount=0, encoding='utf8')
            file_handler.setLevel(Logger.__log_level)
            file_handler.setFormatter(formatter)
            file_handler.suffix = "%Y%m%d%H%M.log"
            Logger.__loggers[tag].addHandler(file_handler)
            if Logger.__use_stdout:
                stream_headler = logging.StreamHandler()
                stream_headler.setLevel(Logger.__log_level)
                stream_headler.setFormatter(formatter)
                Logger.__loggers[tag].addHandler(stream_headler)
        return Logger.__loggers[tag]
