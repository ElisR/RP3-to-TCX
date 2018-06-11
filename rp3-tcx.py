#!/usr/bin/env python

"""
This script converts an RP3 dynamic rower CSV 
workout file to a Garmin TCX file. The TCX file can be then imported
into applications such as Strava or Garmin Connect.

Simply run the program from the shell as follows:

    ./rp3-tcx.py workout.CSV

This will create a new file of the same name but with a TCX extension
in the same directory as the CSV file, i.e. workout.tcx

Set the UTC start time of the workout by using the -t flag
along with the ISO date and time as follows:

    ./rp3-tcx.py [-t 2018-05-14_15:30:00] workout.CSV

Otherwise, the system time is used as the start time.
(This is necessary because RP3 doesn't include the time in the csv file.)

This is very much inspired by Thomas O'Dowd's work with Lemond trainers.
    https://github.com/tpodowd/lemondcsv

"""

# TODO: Handle intervals gracefully

import os
import sys
import csv
import time
import getopt
from xml.etree import ElementTree
from xml.etree.ElementTree import Element, SubElement


XSI = 'http://www.w3.org/2001/XMLSchema-instance'
XSD = 'http://www.garmin.com/xmlschemas/TrainingCenterDatabasev2.xsd'
XML_NS = 'http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2'
EXT_NS = 'http://www.garmin.com/xmlschemas/ActivityExtension/v2'


class Stroke:
    """
    The RP3 Dynamic Rower logs a data point every stroke with information
    such as pace, distance, heartrate, power etc. This object
    represents one particular data point.
    """
    def __init__(self, csvrow):
        self.secs = float(csvrow[7])
        self.speed = self.paceToSpeed(float(csvrow[11]))
        self.dist = float(csvrow[9])
        self.power = float(csvrow[4])
        self.heart = int(csvrow[14])
        self.cadence = int(float(csvrow[6]))
        self.calories = int(self.kJToCalories(float(csvrow[13])))
        self.interval_id = int(csvrow[1])

        def __str__(self):
            return "%d %f %d" % (self.secs, self.speed, self.power)

    def paceToSpeed(self, pace):
        # Converting the time per 500m to m/s
        meters_per_sec = 0
        if pace > 0:
            meters_per_sec = 500 / pace

        return meters_per_sec

    def kJToCalories(self, kJ):
        return kJ / 4.184

    def getIntervalID(self):
        return self.interval_id

    def trackpointElement(self, start):
        tp = Element('Trackpoint')
        time = SubElement(tp, 'Time')
        time.text = Workout.isoTimestamp(start + self.secs)
        dist = SubElement(tp, 'DistanceMeters')
        dist.text = str(self.dist)
        heart = SubElement(tp, 'HeartRateBpm')
        heartvalue = SubElement(heart, 'Value')
        heartvalue.text = str(self.heart)
        cadence = SubElement(tp, 'Cadence')
        cadence.text = str(self.cadence)
        ext = SubElement(tp, 'Extensions')
        self.trackpointExtension(ext, 'Watts', self.power)
        return tp

    def trackpointExtension(self, ext, tag, text):
        tpx = SubElement(ext, 'TPX', {'xmlns': EXT_NS})
        value = SubElement(tpx, tag)
        value.text = str(text)

    @staticmethod
    def parseStrokeHdr(csvrow):
        """
        We assume the order of the fields when parsing the points
        so we fail if the headers are unexpectedly ordered or
        missing or more than expected.
        """
        if len(csvrow) != 25:
            raise Exception("Expected 25 cols, got %d" % len(csvrow))
        exp = []

        exp.append("id")
        exp.append("workout_interval_id")
        exp.append("ref")
        exp.append("stroke_number")
        exp.append("power")
        exp.append("avg_power")
        exp.append("stroke_rate")
        exp.append("time")
        exp.append("stroke_length")
        exp.append("distance")
        exp.append("distance_per_stroke")
        exp.append("estimated_500m_time")
        exp.append("energy_per_stroke")
        exp.append("energy_sum")
        exp.append("pulse")
        exp.append("work_per_pulse")
        exp.append("peak_force")
        exp.append("peak_force_pos")
        exp.append("rel_peak_force_pos")
        exp.append("drive_time")
        exp.append("recover_time")
        exp.append("k")
        exp.append("curve_data")
        exp.append("stroke_number_in_interval")
        exp.append("avg_calculated_power")

        if exp != csvrow:
            raise Exception("Unexpected Header %s != %s" % (exp, csvrow))


class Interval:
    """
    A single workout may be split into different intervals on the rower.
    This object represents a single interval, though the CSV doesn't
    distinguish between resting and active intervals so they are all treated equally here.
    """

    def __init__(self, ID, lap_startsec):
        self.startsec = lap_startsec
        self.endsec = lap_startsec

        self.maxSpeed = 0
        self.maxHeart = 0
        self.maxCadence = 0
        self.maxWatts = 0
        self.ttlDist = 0

        self.interval_id = ID

        self.points = []

    def addStroke(self, s):
        self.points.append(s)
        self.collectStats(s)

    def collectStats(self, p):
        self.endsec = self.startsec + p.secs

        if p.speed > self.maxSpeed:
            self.maxSpeed = p.speed
        if p.heart > self.maxHeart:
            self.maxHeart = p.heart
        if p.cadence > self.maxCadence:
            self.maxCadence = p.cadence
        if p.power > self.maxWatts:
            self.maxWatts = p.power

    def getIntervalID(self):
        return self.interval_id

    def addLap(self, act):
        st = Workout.isoTimestamp(self.startsec)
        lap = SubElement(act, 'Lap', {'StartTime': st})
        last = len(self.points) - 1
        tts = SubElement(lap, 'TotalTimeSeconds')
        tts.text = str(self.points[last].secs)
        dist = SubElement(lap, 'DistanceMeters')
        dist.text = str(self.points[last].dist)
        ms = SubElement(lap, 'MaximumSpeed')
        ms.text = str(self.maxSpeed)
        calories = SubElement(lap, 'Calories')
        calories.text = str(self.points[last].calories)
        maxheart = SubElement(lap, 'MaximumHeartRateBpm')
        maxheartvalue = SubElement(maxheart, 'Value')
        maxheartvalue.text = str(self.maxHeart)
        intensity = SubElement(lap, 'Intensity')
        intensity.text = 'Active'
        trigger = SubElement(lap, 'TriggerMethod')
        trigger.text = 'Manual'
        lap.append(self.trackElement())

    def LapExtension(self, ext, tag, text):
        tpx = SubElement(ext, 'LX', {'xmlns': EXT_NS})
        value = SubElement(tpx, tag)
        value.text = str(text)

    def trackElement(self):
        t = Element('Track')
        for p in self.points:
            t.append(p.trackpointElement(self.startsec))
        return t


class Workout:
    """
    The object represents the complete RP3 workout file.
    """
    def __init__(self, file, start_time):
        self.intervals = []
        self.startsec = time.mktime(start_time)

        self.readCSV(file)

    def readCSV(self, file):
        fp = open(file, 'rt')
        rdr = csv.reader(fp)
        Stroke.parseStrokeHdr(next(rdr))
        for row in rdr:
            p = Stroke(row)
            ID = p.getIntervalID()

            interval = None
            current_ID = 0
            end_time = self.startsec
            if self.intervals:
                interval = self.intervals[-1]
                current_ID = interval.getIntervalID()
                end_time = interval.endsec

            if (ID > current_ID):
                interval = Interval(ID, end_time)
                self.intervals.append(interval)

            interval.addStroke(p)

    @staticmethod
    def isoTimestamp(seconds):
        # Use UTC for isoTimestamp
        tm = time.gmtime(seconds)
        return time.strftime("%Y-%m-%dT%H:%M:%S.000Z", tm)

    def writeTCX(self, file):
        tcdb = self.trainingCenterDB()
        et = ElementTree.ElementTree(tcdb)
        try:
            et.write(file, 'UTF-8', True)
        except TypeError:
            # pre-python 2.7
            et.write(file, 'UTF-8')

    def trainingCenterDB(self):
        dict = {'xsi:schemaLocation': XML_NS + ' ' + XSD,
                'xmlns': XML_NS,
                'xmlns:xsi': XSI}
        tcdb = Element('TrainingCenterDatabase', dict)
        acts = SubElement(tcdb, 'Activities')
        self.addActivity(acts)
        self.addAuthor(tcdb)
        return tcdb

    def addActivity(self, acts):
        act = SubElement(acts, 'Activity', {'Sport': 'Rowing'})
        id = SubElement(act, 'Id')
        id.text = Workout.isoTimestamp(self.startsec)
        for interval in self.intervals:
            interval.addLap(act)

        self.addCreator(act)

    def addCreator(self, act):
        c = SubElement(act, 'Creator', {'xsi:type': 'Device_t'})
        name = SubElement(c, 'Name')
        name.text = 'RP3 Dynamic Rower'
        unit = SubElement(c, 'UnitId')
        unit.text = '0'
        prd = SubElement(c, 'ProductID')
        prd.text = '0'
        ver = SubElement(c, 'Version')
        vmaj = SubElement(ver, 'VersionMajor')
        vmaj.text = '1'
        vmin = SubElement(ver, 'VersionMinor')
        vmin.text = '0'
        bmaj = SubElement(ver, 'BuildMajor')
        bmaj.text = '0'
        bmin = SubElement(ver, 'BuildMinor')
        bmin.text = '0'

    def addAuthor(self, tcdb):
        a = SubElement(tcdb, 'Author', {'xsi:type': 'Application_t'})
        name = SubElement(a, 'Name')
        name.text = 'RP3 CSV to TCX Convertor'
        build = SubElement(a, 'Build')
        ver = SubElement(build, 'Version')
        vmaj = SubElement(ver, 'VersionMajor')
        vmaj.text = '1'
        vmin = SubElement(ver, 'VersionMinor')
        vmin.text = '0'
        bmaj = SubElement(ver, 'BuildMajor')
        bmaj.text = '0'
        bmin = SubElement(ver, 'BuildMinor')
        bmin.text = '0'
        lang = SubElement(a, 'LangID')
        lang.text = 'en'
        partnum = SubElement(a, 'PartNumber')
        partnum.text = 'none'


def output_name(iname):
    # Validate name ends with .CSV
    if iname.lower().endswith(".csv"):
        prefix = iname[:-3]
        oname = prefix + "tcx"
        if not os.path.exists(oname):
            return oname
        else:
            raise Exception("File %s already exists. Cannot continue." % oname)
    else:
        raise Exception("%s does not end with .csv" % iname)


def usage_exit():
    sys.stderr.write("Usage: rp3-tcx.py [-f workout.tcx -t yyyy-mm-dd_hh:MM:ss] workout.csv\n")
    sys.exit(1)

opts, args = getopt.getopt(sys.argv[1:], 'f:t:h')
oname = None
workout_start = time.gmtime()
for opt, arg in opts:
    if opt == '-f':
        oname = arg
    elif opt == '-t':
        workout_start = time.strptime(arg, "%Y-%m-%d_%H:%M:%S")
    elif opt == '-h':
        usage_exit()

if len(args) != 1:
    usage_exit()
else:
    iname = args[0]
    if oname is None:
        oname = output_name(iname)
    work = Workout(iname, workout_start)
    if oname == '-':
        if hasattr(sys.stdout, 'buffer'):
            ofile = sys.stdout.buffer
        else:
            ofile = sys.stdout
    else:
        sys.stderr.write("Writing to: %s\n" % oname)
        ofile = open(oname, "wb")
    work.writeTCX(ofile)
    if oname != '-':
        ofile.close()