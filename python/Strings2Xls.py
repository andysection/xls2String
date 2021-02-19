# -*- coding:utf-8 -*-

import os
from optparse import OptionParser
from StringsFileUtil import StringsFileUtil
import pyExcelerator
import time

# Add command option


def addParser():
    parser = OptionParser()

    parser.add_option("-f", "--stringsDir",
                      help=".strings files directory.",
                      metavar="stringsDir")

    parser.add_option("-t", "--targetDir",
                      help="The directory where the excel(.xls) files will be saved.",
                      metavar="targetDir")

    (options, _) = parser.parse_args()

    return options

#  convert .strings files to single xls file


def convertToSingleFile(stringsDir, targetDir):
    destDir = targetDir + "/strings-files-to-xls_" + \
        time.strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(destDir):
        os.makedirs(destDir)

    # Create xls sheet
    for _, dirnames, _ in os.walk(stringsDir):
        for _, _, filenames in os.walk(stringsDir):
            stringsFiles = [
                fi for fi in filenames if fi.endswith(".strings")]
            for stringfile in stringsFiles:
                fileName = stringfile.replace(".strings", "")
                filePath = destDir + "/" + fileName + ".xls"
                if not os.path.exists(filePath):
                    workbook = pyExcelerator.Workbook()
                    ws = workbook.add_sheet(fileName)
                    index = 0
                    if index == 0:
                        ws.write(0, 0, 'keyName')
                    ws.write(0, index+1, fileName)

                    path = stringsDir+ '/' + stringfile
                    (keys, values) = StringsFileUtil.getKeysAndValues(
                        path)
                    for x in range(len(keys)):
                        key = keys[x]
                        value = values[x]
                        if (index == 0):
                            ws.write(x+1, 0, key)
                            ws.write(x+1, 1, value)
                        else:
                            ws.write(x+1, index + 1, value)
                    index += 1
                    workbook.save(filePath)

    print "Convert %s successfully! you can see xls file in %s" % (
        stringsDir, destDir)


def startConvert(options):
    stringsDir = options.stringsDir
    targetDir = options.targetDir

    print "Start converting"

    if stringsDir is None:
        print ".strings files directory can not be empty! try -h for help."
        return

    if targetDir is None:
        print "Target file directory can not be empty! try -h for help."
        return

    convertToSingleFile(stringsDir, targetDir)
    


def main():
    options = addParser()
    startConvert(options)

main()
