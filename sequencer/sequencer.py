#!/usr/bin/env python3
#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
 pyblaze-sequencer

 A library that reads a Google Sheet and sends timed commands to Pixelblazes
 on the local network.
"""


# ----------------------------------------------------------------------------

__version__ = "0.0.1"

#| Version | Date       | Author        | Comment                                 |
#|---------|------------|---------------|-----------------------------------------|
#|  v0.0.1 | 12/11/2022 | jvyduna       | Created |

# ----------------------------------------------------------------------------


from pixelblaze import *
from playsound import playsound
import gspread
import time
import re
#from urllib.parse import parse_qsl

# import curses
# import pydub


class Sequencer:
    UPDATE_RATE = 60.0 # loops per second

    interval = 1 / UPDATE_RATE
    devices_patterns = {} # {Pixelblaxe: pattern_name_string[]}
    devices = {} # {device_name: Pixeblaze()}
    device_names = []
    ips = []
    patterns = {} # {pattern_name_string: device_name[] }
    sheetname = ""
    sequence_sheet = None

    def __init__(self, sheetname):
        self.sheetname = sheetname
        self.gc = gspread.oauth()

        self.discover()
        # self.poplate_patterns_sheet()
        self.load_sequence()
        self.load_music()


    def discover(self):
        print("Finding Pixelblazes...")
        # Enumerate the available Pixelblazes on the network.

        syncEnumerator = PixelblazeEnumerator()
        syncEnumerator.enableTimesync()
        time.sleep(1.5) # Time enough to discover all?
        syncEnumerator.stop()
        print(f"    Found {len(syncEnumerator.devices)} device(s) and sync'd time()")
        
        #for ip in Pixelblaze.EnumerateAddresses(timeout=1500):
        for ip in [v['address'][0] for v in syncEnumerator.devices.values()]:
            try:
                pb = Pixelblaze(ip)
                device_name = pb.getDeviceName()
                self.devices[device_name] = pb
                patterns = pb.getPatternList()
                self.devices_patterns[pb] = patterns
                for pattern_id, pattern_name in patterns.items():
                    self.patterns.setdefault(pattern_name, []).append(device_name)
                self.device_names.append(device_name)
                self.ips.append(ip)
                print(f"    at {ip} found '{device_name}' with {len(patterns)} patterns")
                
            except ConnectionResetError:
                print(f"    at {ip} (Device websocket connection error)")
            except Exception as e:
                print(f"    at {ip}, {format(e)}")

            #Speed up testing! TODO remove
            # if len(self.ips) > 1:
            #     break

    

    def poplate_patterns_sheet(self):
        self.devices_patterns_sheet = self.gc.open(self.sheetname).worksheet("Devices")

        # Build table to pupulate Devices / Patterns sheet
        device_patterns_table = []
        self.devices_patterns_sheet.batch_clear(['A1:Z400'])
        
        # Populate device/pattern inventory column headers (devices)
        device_patterns_table.append([None] + self.ips)
        device_patterns_table.append([None] + self.device_names)

        # Print row (pattern) headers from most to least prevalent
        counts = {k: len(v) for k, v in self.patterns.items()}
        counts = dict(sorted(counts.items(), key=lambda item: (-item[1], item[0])))


        # Accumulate row headers (pattern names) and 1 or blank for each device
        for pattern_name in counts.keys():
            row = [pattern_name]
            for device_name in self.device_names:
                row.append(1 if device_name in self.patterns[pattern_name] else None)
            device_patterns_table.append(row)
        
        self.devices_patterns_sheet.update(device_patterns_table)

    def load_sequence_sheet(self):
        if not self.sequence_sheet:
            self.sequence_sheet = self.gc.open(self.sheetname).worksheet("Sequence")

    def load_sequence(self):
        self.load_sequence_sheet()

        # Load list of expected device names
        device_cell_re = re.compile(r'^Devices$')
        device_cell = self.sequence_sheet.find(device_cell_re, case_sensitive=True)
        seq_devices_range = gspread.utils.rowcol_to_a1(device_cell.row+1, device_cell.col+1) + ':' + gspread.utils.rowcol_to_a1(device_cell.row + 1, device_cell.col + 20)
        seq_dev_statuses_range = gspread.utils.rowcol_to_a1(device_cell.row+2, device_cell.col+1) + ':' + gspread.utils.rowcol_to_a1(device_cell.row + 2, device_cell.col + 20)
        self.seq_devices = self.sequence_sheet.get_values(seq_devices_range)[0]
        
        # Update if devices were found        
        found = ['Online' if sd in self.device_names else '' for sd in self.seq_devices]
        self.sequence_sheet.update(seq_dev_statuses_range, [found])

        # Load rows of timestamps and commands (patterns to launch or vars to send)
        self.tc = self.sequence_sheet.find(re.compile(r'^Trigger timestamp$'))
        seq_range = gspread.utils.rowcol_to_a1(self.tc.row+1, self.tc.col) + ':' + gspread.utils.rowcol_to_a1(self.tc.row + 200, device_cell.col + 1 + len(self.seq_devices))
        timecoded_sequence = self.sequence_sheet.get_values(seq_range)
        self.sequence_sheet.format(seq_range, { "textFormat": { "bold": False } })

        self.sequence = []

        seconds = 0
        # ['hh:mm:ss.ms', *cmds] => [seconds_f, *cmds]
        for row in timecoded_sequence:
            timecode = row[0]
            try:
                seconds = [float(x) for x in timecode.split(':')]
                seconds = seconds[0]*3600 + seconds[1]*60 + seconds[2]
            except ValueError as e:
                if re.search("could not convert string to float", str(e)) is None:
                    raise e
                # Otherwise reuse a prior row's sucessful tc parse
            
            self.sequence.append([seconds] + row[1:])

        # could dry-run parsing times and commands here and highlight problems/ignores


    def load_music(self):
        self.load_sequence_sheet()
        self.music_file = "music/" + self.sequence_sheet.get("B1").first()
        print("Music file: " + self.music_file)


    def play_sound(self):
        try:
            playsound(self.music_file, False)
        except:
            playsound("../" + self.music_file, False)

    def run(self):
        self.started_at = time.time() # - self.offset
        self.last_loop_at = time.time()
        self.play_sound()
        self.running = True
        self.cursor = 0    # Index of row in the sequence to look for
        while self.running:
            self.__loop()
            time.sleep(self.interval - ((time.time() - self.started_at) % self.interval))


    def __loop(self):
        self.delta = time.time() - self.last_loop_at
        self.last_loop_at = time.time()
        self.elapsed = time.time() - self.started_at
        
        if self.elapsed > self.sequence[self.cursor][0]: # just passed timecode + offset:
            self.execute_row(self.sequence[self.cursor][1:])
            self.cursor += 1

        # if sound.done or key input = "q" or no more rows or row timecode blank
        if self.cursor >= len(self.sequence):
            self.running = False
            self.cursor = 0
            print("Sequence complete")

    def execute_row(self, commands):
        print(f"   {self.elapsed: .2f} elapsed, executing: {commands}")

        self.tc
        current_tc_cell = gspread.utils.rowcol_to_a1(self.tc.row + self.cursor + 1, self.tc.col)
        self.sequence_sheet.format(current_tc_cell, { "textFormat": { "bold": True } })

        all_command = commands[0]
        # If "All" column has a command, all other device columns are ignored
        if all_command.strip():
            valid_devices = list(set(self.seq_devices) & set(self.device_names))
            dev_count_str = f'{len(valid_devices)} device{"s"[:len(valid_devices)^1]}'
            for device_name in valid_devices:
                pb = self.devices[device_name]
                if self._is_varString(all_command):
                    pattern_vars = self._parse_var_str(all_command)
                    pb.setActiveVariables(pattern_vars)
                elif device_name in self.patterns[all_command]:
                    pb.setActivePatternByName(commands[0])
            if self._is_varString(all_command):
                print(f'      setVars {pattern_vars} to {dev_count_str}')
            else:
                print(f'      Sent "{all_command}" to {dev_count_str}')

        else:
            for i, command in enumerate(commands[1:]):
                if not command.strip(): continue
                device_name = self.seq_devices[i]
                if device_name in self.devices:
                    pb = self.devices[device_name]
                    if self._is_varString(command):
                        pattern_vars = self._parse_var_str(command)
                        pb.setActiveVariables(pattern_vars)
                        print(f'      setVars {pattern_vars} to {device_name}')
                    elif command in self.patterns and device_name in self.patterns[command]:
                        pb.setActivePatternByName(command)
                        print(f'      Activated "{command}" on {device_name}')

    def _parse_var_str(self, str):
        pairs = str.strip().split('&')
        pattern_vars = dict(map(lambda x: x.split("="), pairs))
        return pattern_vars

    def _is_varString(self, str):
        return bool(re.search(r'^[\w]+=[\w.]+', str.strip()))


# Testing...
s = Sequencer("PBS Midnight Blessings")
# s.poplate_patterns_sheet()
s.run()






