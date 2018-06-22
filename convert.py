#!/usr/bin/env python

""" Converts XLSX files to CSV

    Usage:
    python convert.py [source] [output format]
"""

import argparse
import csv
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderQuotaExceeded, GeocoderUnavailable
import json
import openpyxl

parser = argparse.ArgumentParser(description='Convert XLSX files to CSV')
parser.add_argument('source', type=str, nargs=1,
                    help='the source XLSX file to be converted')
parser.add_argument('-o', '--output', default='all', choices=['csv', 'json', 'geojson'],
                    help='output format', )
args = parser.parse_args()


class SourceFile():
    def __init__(self, filepath):
        self.filepath = filepath

    def load_data(self):
        try:
            wb = openpyxl.load_workbook(self.filepath)
            data = wb.active
            return data
        except Exception as e:
            print "Error loading data: {}".format(e)
            return False


class DestinationFile():
    def __init__(self, data, source_filepath):
        self.data = data
        self.source_filepath = source_filepath

    def get_filepath(self, extension):
        filepath = self.source_filepath.replace('.xlsx', extension)
        return filepath

    def get_coordinates(self, data):
        if (data[15].value and data[16].value):
            return [unicode(data[16].value), unicode(data[15].value)]
        else:
            geolocator = Nominatim()
            zip = "{}-{}".format(data[10].value, data[11].value) if data[11].value else data[10].value
            print "Geocoding {}".format(data[7].value)
            try:
                location = geolocator.geocode("{} {} {} {}".format(data[7].value, data[9].value, data[13].value, zip), timeout=7)
                if location:
                    return [location.longitude, location.latitude]
            except GeocoderTimedOut:
                self.get_coordinates(data)
            except GeocoderQuotaExceeded:
                print "Geolocation service quota exceeded"
                return False
            except GeocoderUnavailable:
                print "Geolocation service unavailable"
                return False

    def write_csv(self):
        try:
            destination = self.get_filepath('.csv')
            with open(destination, 'wb') as f:
                c = csv.writer(f)
                for r in self.data.rows:
                    new_row = []
                    for cell in r:
                        value = unicode(cell.value) if cell.value else ''
                        new_row.append(value.encode('utf-8'))
                    c.writerow(new_row)
            print "CSV file created at {}".format(destination)
        except Exception as e:
            print "Error creating CSV file: {}".format(e)

    def write_json(self):
        try:
            destination = self.get_filepath('.json')
            rows = list(data)
            array = []
            for x in range(1, len(rows)):
                part = {}
                for n in range(0, 23):
                    part[rows[0][n].value.replace("*", "")] = unicode(rows[x][n].value)
                array.append(part)
            with open(destination, 'wb') as f:
                f.write(json.dumps(array, sort_keys=True, indent=4,
                        separators=(',', ': ')))
            print "JSON file created at {}".format(destination)
        except Exception as e:
            print "Error creating JSON file: {}".format(e)

    def write_geojson(self):
        try:
            destination = self.get_filepath('-geo.json')
            array = []
            rows = list(data)
            for x in range(1, len(rows)):
                part = {"type": "Feature",
                        "geometry": {
                            "type": "Point",
                            "coordinates": self.get_coordinates(rows[x])
                            },
                        "properties": {
                            "name": unicode(rows[x][0].value),
                            "type": unicode(rows[x][5].value),
                            "url": unicode(rows[x][14].value)
                            }
                        }
                array.append(part)
            with open(destination, 'wb') as f:
                f.write(json.dumps({"type": "FeatureCollection", "features": array},
                        indent=4, separators=(',', ': ')))
            print "geoJSON file created at {}".format(destination)
        except Exception as e:
            print "Error creating geoJSON file: {}".format(e)

    def write_all(self):
        self.write_csv()
        self.write_json()
        self.write_geojson()


source = SourceFile(args.source[0])
data = source.load_data()
destination = DestinationFile(data, source.filepath)
getattr(destination, 'write_{}'.format(args.output))()
