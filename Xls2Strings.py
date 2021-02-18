# -*- coding:utf-8 -*-

from optparse import OptionParser
from XlsFileUtil import XlsFileUtil
from StringsFileUtil import StringsFileUtil
from Log import Log
import os
import time


def addParser():
    parser = OptionParser()

    parser.add_option("-f", "--fileDir",
                      help="Xls files directory.",
                      metavar="fileDir")

    parser.add_option("-t", "--targetDir",
                      help="The directory where the strings files will be saved.",
                      metavar="targetDir")

    parser.add_option("-a", "--additional",
                      help="additional info.",
                      metavar="additional")

    (options, args) = parser.parse_args()
    Log.info("options: %s, args: %s" % (options, args))

    return options


def convertFromSingleForm(options, fileDir, targetDir):
    for _, _, filenames in os.walk(fileDir):
        xlsFilenames = [fi for fi in filenames if fi.endswith(".xls")]
        for file in xlsFilenames:
            xlsFileUtil = XlsFileUtil(fileDir+"/"+file)
            table = xlsFileUtil.getTableByIndex(0)
            firstRow = table.row_values(0)
            keys = table.col_values(0)
            del keys[0]

            for index in range(len(firstRow)):
                if index <= 0:
                    continue
                languageName = firstRow[index]
                values = table.col_values(index)
                del values[0]
                StringsFileUtil.writeToFile(
                    keys, values, targetDir + "/", languageName+".strings", options.additional)
    print "Convert %s successfully! you can see strings file in %s" % (
        fileDir, targetDir)


def startConvert(options):
    fileDir = options.fileDir
    targetDir = options.targetDir

    print "Start converting"

    if fileDir is None:
        print "xls files directory can not be empty! try -h for help."
        return

    if targetDir is None:
        print "Target file directory can not be empty! try -h for help."
        return

    targetDir = targetDir + "/xls-files-to-strings_" + \
        time.strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(targetDir):
        os.makedirs(targetDir)

    convertFromSingleForm(options, fileDir, targetDir)


def main():
    options = addParser()
    startConvert(options)


main()
