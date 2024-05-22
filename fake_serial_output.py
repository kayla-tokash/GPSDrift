

#requires socat
#todo, replace this with an all python solution

from random import randint, gauss
import serial
import os
import time
import threading
import pty
from rocket_data import DataPacket, WrappedHeader
from json import JSONDecoder, dumps as json_pretty
import ctypes

#OPEN_FAKE_SERIAL_CMD = "socat -d -d pty,raw,echo=0 pty,raw,echo=0"

"""
2022/12/23 20:25:47 socat[40094] N PTY is /dev/ttys004
2022/12/23 20:25:47 socat[40094] N PTY is /dev/ttys005
2022/12/23 20:25:47 socat[40094] N starting data transfer loop with FDs [5,5] and [7,7]
"""


sout = None
sout_lock = threading.Lock()
def set_serial_out(sout_new, overwrite=True, l=sout_lock):
    l.acquire(blocking=True)
    global sout
    if sout != None:
        if overwrite:
            sout.close()
            sout = None
        else:
            return
    sout = sout_new
    l.release()

def write_sout(data, s=sout, l=sout_lock):
    l.acquire(blocking=True)
    s.write(data)
    l.release()

sin = None
sin_lock = threading.Lock()
def set_serial_in(sin_new, overwrite=True, l=sin_lock):
    l.acquire(blocking=True)
    global sin
    if sin != None:
        if overwrite:
            sin.close()
            sin = None
        else:
            return
    sin = sin_new
    l.release()

def read_device(s=sin, l=sin_lock, num_bytes=None, callback=None, loop=True):
    callback_result = True
    l.acquire(blocking=True)
    while callback_result and loop:
        data = s.read(num_bytes)
        if not callback is None:
            callback_result = callback(data)
    l.release()

def create_serial_devices(gen_sin=True, gen_sout=True):
    pty_in, pty_out = pty.openpty()

    print("pty_in =")
    print(pty_in)

    #sin_name = os.ttyname(pty_in)
    sout_name = os.ttyname(pty_out)

    #print("sin_name = %s,\nsout_name = %s" %  (sin_name, sout_name))

    sout = None
    sin = None
    
    if gen_sout:
        sout = serial.Serial(sout_name, 9600, timeout=1)
    if gen_sin:
        sin = pty_in #serial.Serial(sin_name, 9600, timeout=1)

    return (sin, sout)

start_time=int(time.time()*1000)
def create_packet(runtime=-1, accel=(0,0,0), gyro=(0,0,0), magno=(0,0,0), alt=0, temperature=0):
        if runtime < 0:
            runtime = int(time.time()*1000) - start_time
        packet = DataPacket(json={
            "runtime":runtime,
            "accelerometer":accel,
            "gyroscope":gyro,
            "magnometer":magno,
            "altitude":alt,
            "temperature":temperature,
            "status":0
        })
        return packet.serialize()

            # newtons

# https://www.questaerospace.com/products/product_db2a587b-e6c9-0d42-e6dd-af147938dd8f
AEROTECH_F44_8W= {
    "impulse":41.5,
    "average_thrust":44,
    "max_thrust":50,
    "burn_time":1000, #ms
    "launch_mass":48+19.7, #grams
    "empty_mass":48, #grams
    "delay":8000 #ms
}

ESTES_F15 = {
    "impulse":50,
    "average_thrust":15.0,
    "max_thrust":25.62,
    "burn_time":3450, #ms
    "launch_mass":60, #grams
    "empty_mass":30, #grams
    "delay":4000 #ms
}

AT_G138T = {
    "impulse":167,
    "average_thrust":153,
    "max_thrust":195,
    "burn_time":1090, #ms
    "launch_mass":152, #grams
    "empty_mass":81.9, #grams
    "delay":14000 #ms
}

AT_G64 = {
    "impulse":119,
    "average_thrust":62.9,
    "max_thrust":98.3,
    "burn_time":1880, #ms
    "launch_mass":151, #grams
    "empty_mass":88.7, #grams
    "delay":10000 #ms
}


AT_F52T = {
    "name":"Aerotech F52 Blue Thunder",
    "impulse":73,
    "average_thrust":53.6,
    "max_thrust":79,
    "burn_time":1360,
    "launch_mass":121,
    "empty_mass":84.8,
    "delay":10000
}

AT_E16W = {
    "impulse":37.7,
    "average_thrust":20.8,
    "max_thrust":37.2,
    "burn_time":1800, #ms
    "launch_mass":107, #grams
    "empty_mass":88, #grams
    "delay":4000 #ms
}

CUSTOM_O1300 = {
    "impulse":544,
    "average_thrust":1300,
    "max_thrust":195,
    "burn_time":3000, #ms
    "launch_mass":300, #grams
    "empty_mass":100, #grams
    "delay":16000 #ms
}

CUSTOM_I320 = {
    "impulse":544,
    "average_thrust":320,
    "max_thrust":195,
    "burn_time":1690, #ms
    "launch_mass":300, #grams
    "empty_mass":100, #grams
    "delay":16000 #ms
}

ESTES_D12 = {
    "impulse":16.8,
    "average_thrust":10.4,
    "max_thrust":29.7,
    "burn_time":1610, #ms
    "launch_mass":42.6, #grams
    "empty_mass":21.5, #grams
    "delay":5000 #ms
}

AT_I115 = {
    "name":"Aerotech I115-14 White Lightning",
    "impulse":409.3,
    "thrust_profile":[(0,0),(100,120),(500,147),(1000,155),(1300,160),(1500,156),(3200,67),(3400,22.2),(3600,5),(3700,0)],
    "average_thrust":115,
    "max_thrust":172.9,
    "burn_time":3800, #ms
    "launch_mass":580, #grams
    "empty_mass":351, #grams
    "delay":10000 #ms
}

AT_L1090 = {
    "name":"AEROTECH L1090W-PS RMS-54/2800",
    "impulse":2763,
    "thrust_profile":[(0,0),(50,4.45*280),(100,4.45*300),(200,4.45*275),(250,4.45*280),(2000,4.45*237),(2250,4.45*125),(2700,4.45*10),(3000,0)],
    "average_thrust":1090,
    "max_thrust":1372,
    "burn_time":3000,
    "launch_mass":2370,
    "empty_mass":1032,
    "delay":17500 # it's actually plugged
}

#Zarya water rocket
ZARYA_AIR = {
        "impulse":2814,
        "average_thrust":262,
        "max_thrust":846,
        "burn_time":10000,
        "delay":13000,
        "empty_mass":(12.58-7.94)*1000,
        "launch_mass":0
}

AT_I175 = {
    "name":"AT I175-11 (13A) Super White Lightning DMS-38",
    "impulse":333,
    "average_thrust":175,
    "max_thrust":260,
    "delay":11000,
    "burn_time":1900,
    "empty_mass":348-168,
    "launch_mass":348,
    "thrust_profile":[(0,0),(100,245),(250,260),(1750,115),(1850,25),(1900,0)]
}

#default_engine=ESTES_D12
#default_engine=ZARYA_AIR
#default_engine=AT_G64
#default_engine=AT_G138T
#default_engine=CUSTOM_O1300
#default_engine=CUSTOM_IXXT
#default_engine=AT_F52T
#default_engine=AT_I115
#default_engine=AT_L1090
default_engine=AT_I175
#default_engine=ESTES_F15
#default_engine=AEROTECH_F44_8W

#"drag":0.47 *(3.14*pow(2.8/100,2)), # C_d * CrossSectional_Area # yam2

ezi65 = {
    "mass":2128,
    "C_d":0.440,
    "area":3.14 * pow(10.16/100,2)
}

#rocket_mass=520 #yam2
super_sonic_rocket = {
    "mass":94.8,
    "C_d":0.309,
    "area":3.14 * pow(2.56/100,2)
}

yam2_rocket = {
    "mass":615,
    "C_d":0.47,
    "area":3.14 * pow(2.8/100,2)
}

zarya_rocket = {
    "mass":7.94*1000,
    "C_d":0.5,
    "area":183.8538561
}

loc_4in_v2 = {
    "mass":1211.8,
    "C_d":0.24,
    "area":3.14 * pow(2.8/100,2)
}

sim_data = None
def reset_sim(launch_altitude, engine_profile, rocket_data=loc_4in_v2):
    global sim_data
    rocket_mass = rocket_data["mass"]
    sim_data = {
        "launch_altitude":launch_altitude,
        "altitude":launch_altitude,
        "velocity":0,
        "burn_time":engine_profile["burn_time"],
        "burn_mass_rate":(engine_profile["launch_mass"] - engine_profile["empty_mass"]) / engine_profile["burn_time"],
        "engine_thrust":engine_profile["average_thrust"],
        "engine_name":engine_profile["name"] if "name" in engine_profile else None,
        "thrust_profile":engine_profile["thrust_profile"] if "thrust_profile" in engine_profile else None,
        "current_mass":rocket_mass+engine_profile["launch_mass"],
        "first_chute_delay":engine_profile["delay"] + engine_profile["burn_time"],
        "first_fallrate":-27, #yam2 is -10
        "second_chute_altitude":100, #TODO make this configurable
        "second_fallrate":-6.55, #yam2 is -4.5
        "time_elapsed":0,
        "drag":0.47 *(3.14*pow(2.8/100,2)), # C_d * CrossSectional_Area # yam2
        #"drag":0.53 *(3.14*pow(2.1/100,2)), # C_d * CrossSectional_Area # yam1
        "gravity":9.81,
        "engine_deficiency":0.99, #How much of the thrust we get in the real world-- I find it tends to be 0.93-0.95
        "acceleration":-1
    }


temperature = 30 #c
def calc_altitude_sim(time, launch_time=5, engine_details=AT_F52T, randomness=0):
    global sim_data, temperature
    t_seconds = time / 1000
    t_post_launch = time - launch_time*1000
    t_snapshot = time - sim_data["time_elapsed"]
    a_relative = sim_data["altitude"] - sim_data["launch_altitude"]

    sim_data["time_elapsed"] = time
    if t_seconds < launch_time:
        if randomness:
            # this is broken now
            return gauss(sim_data["launch_altitude"], randomness)
        return sim_data

    if not sim_data["engine_name"] is None:
        print(sim_data["engine_name"])

    # engine thrust is in Newtons (kg*m/s^2) but we're using grams and ms lol
    engine_thrust = sim_data["engine_thrust"] # default to the average thrust
    # calculate current thrust
    if not sim_data["thrust_profile"] is None:
        for index in range(len(sim_data["thrust_profile"])):
            if t_post_launch < sim_data["thrust_profile"][index][0]:
                if sim_data["thrust_profile"][index][1] == 0:
                    engine_thrust = 0
                    break
                if index == 0:
                    d_thrust = sim_data["thrust_profile"][0][1]
                    engine_thrust = d_thrust * (t_post_launch / sim_data["thrust_profile"][0][0])
                else:
                    d_thrust = sim_data["thrust_profile"][index][1] - sim_data["thrust_profile"][index-1][1]
                    time_ratio = t_post_launch / (sim_data["thrust_profile"][index][0] - sim_data["thrust_profile"][index-1][0])
                    engine_thrust = d_thrust * time_ratio
                break
        engine_thrust = 0 #end of thrust curve

    distance = lambda t,v:(t/1000) * v
    engine_acceleration = sim_data["engine_thrust"] / (sim_data["current_mass"] / 1000)
    velocity = lambda t:sim_data["engine_deficiency"] * (t/1000) * engine_acceleration
    coast = lambda t:(t/1000) * (-sim_data["gravity"])
    drag_acceleration = 0.5 * 1.225 * pow(sim_data["velocity"], 2) * sim_data["drag"]
    drag = lambda t:-(t/1000) * drag_acceleration
    d_velocity = coast(t_snapshot) + drag(t_snapshot)
    print("drag_acc:" + str(drag_acceleration))
    print("engine_acc:" + str(engine_acceleration))
    if sim_data["burn_time"] > 0:
        if sim_data["burn_time"] > t_snapshot:
            sim_data["burn_time"] -= t_snapshot
            d_velocity += velocity(t_snapshot)
            sim_data["current_mass"] -= sim_data["burn_mass_rate"]
            print("Burning")
        elif sim_data["burn_time"] > 0:
            sim_data["current_mass"] -= sim_data["burn_mass_rate"] * (sim_data["burn_time"] / t_snapshot)
            d_velocity += velocity(sim_data["burn_time"])
            sim_data["burn_time"] = 0
            print("Burn complete")

    sim_data["acceleration"] = ((-engine_acceleration if sim_data["burn_time"] > 0 else 0) + drag_acceleration + (0 if a_relative > 0 else -sim_data["gravity"])) / sim_data["gravity"]

    if sim_data["velocity"] > 340:
        temperature += 5
    else:
        temperature = max(20, temperature - .1)

    if t_post_launch < sim_data["first_chute_delay"]:
        print("FLYING")
        sim_data["velocity"] += d_velocity
    elif a_relative > sim_data["second_chute_altitude"]:
        print("FIRST CHUTE")
        sim_data["velocity"] = sim_data["first_fallrate"]
        sim_data["acceleration"] = -1
    elif sim_data["altitude"] > sim_data["launch_altitude"]:
        print("SECOND CHUTE")
        sim_data["velocity"] = sim_data["second_fallrate"]
        sim_data["acceleration"] = -1

    d_distance = distance(t_snapshot, sim_data["velocity"])
    if sim_data["altitude"] > sim_data["launch_altitude"] or d_distance + sim_data["altitude"] > sim_data["launch_altitude"]:
        sim_data["altitude"] += d_distance
    else:
        sim_data["altitude"] = sim_data["launch_altitude"]

    print("time: %f, launch: %f, snapshot: %f"  % (t_seconds, t_post_launch, t_snapshot))
    print("velocity: %f, d_vel: %f"  % (sim_data["velocity"], d_velocity))
    print("altitude: %f, d_alt: %f"  % (sim_data["altitude"], d_distance))

    if randomness:
        return gauss(sim_data["altitude"], randomness)

    print(json_pretty(sim_data, indent=4))

    return sim_data

def calc_altitude_algerbraic(time, launch_time=5, time_to_apogee=14.30, deploy_chute_time=32.00, flight_time=55.00, apogee=1308.0, launch_altitude=20, randomness=0, thrust=None):
    t_seconds = time / 1000.
    #if t_seconds < launch_time or t_seconds > flight_time:
    #    return 0
    print("t_seconds: %f, time: %f, launch_time: %f, time_to_apogee: %f, flight_time: %f, apogee: %f" % (t_seconds, time, launch_time, time_to_apogee, flight_time, apogee))
    random_offset = randint(0, randomness) / 100
    time_at_apogee = time_to_apogee + launch_time
    if t_seconds > launch_time and t_seconds < time_at_apogee:
        print("in flight")
        # y = (x+a)*(-x+b) = -x^2 + -(a + b)x + ab
        # c = a + b, d = a*b
        # apogee = (time_to_apogee+a)*(-time_to_apogee+b)
        # time = 0; y = 0; 0 = a
        # time = time_to_apogee; y = apogee; a = 0; apogee = (time_to_apogee+0)*(-time_to_apogee+b)
        #        (apogee/ttime_to_apogee)+time_to_apogee
        # d_y(t=0) > 0, d_y(t=time_to_apogee) = 0; d_y(t) = e + fx where b < 0
        # a = -1->f = -2 using derivative; so 0 = e - 2 * time_to_apogee; then e = 2 * time_to_apogee
        #
        # y(0) = 0; y(time_to_apogee) = apogee; y(2*time_to_apogee) = 0
        # y(t) = a * x^2 + b * x + c; a = -1; b = e = (2 * time_to_apogee); c = ?
        # apogee = - (time_to_apogee) ^ 2 + (2 * time_to_apogee) * time_to_apogee + c
        # c = apogee + (time_to_apogee) ^ 2 - (2 * (time_to_apogee ^ 2))
        # y = (x+a)*(x+b)
        # apogee = (toa+a)*(toa-b)
        # time = 0; y = 0; a = 0
        # time = toa; y = apogee; a= 0; apogee=(toa+0)*(toa-b)
        #       b = toa - apogee/toa
        #
        a = 0
        b = (apogee / time_to_apogee) + time_to_apogee
        c = 0.0
        formula = "(x+a)*(-x+b)"
        #                                              -(t-launch_time)
        altitude = lambda t: (t + a) * (-t + b)
        #a = -1
        #b = 2 * time_to_apogee
        #c = apogee - pow(time_to_apogee, 2)
        #formula = "a*x^2 + b*x + c"
        #altitude = lambda t: a * pow(t, 2) + b * t + c
        print("a=%f, b=%f, c=%f, formula=%s, results=%f" % (a, b, c, formula, altitude(t_seconds - launch_time)))
        return altitude(t_seconds - launch_time) + launch_altitude + random_offset
    if t_seconds <= launch_time or t_seconds > launch_time + flight_time:
        print("on gouund")
        return launch_altitude + random_offset
    #if t_seconds > launch_time and t_seconds > time_at_apogee:
    
    print("falling")
    fallrate = apogee / (flight_time - time_to_apogee)
    altitude = apogee - fallrate * (t_seconds - time_at_apogee)
    return altitude  + launch_altitude + random_offset

imported_data = None
def load_data(filename="data.json", force_reload=False):
    global imported_data
    if imported_data is None or force_reload:
        with open(filename, "r") as file_in:
            imported_data = JSONDecoder().decode("".join(file_in.readlines()))
            if "magno_calibration" in imported_data.keys():
                imported_data.pop("magno_calibration")
    return imported_data

def use_data(time, json_data=None):
    if json_data is None:
        json_data = load_data()
    data_start_time = json_data["runtime"][0]
    # seek the nearest time recorded
    current_datapoint={}
    current_time_index = 0
    while (json_data["runtime"][current_time_index] - data_start_time) < time:
        current_time_index+=1
        if current_time_index >= len(json_data["runtime"]):
            # return the last value and because we have no more data
            for key in json_data.keys():
                current_datapoint[key]=json_data[key][-1]
            current_datapoint["runtime"] = int(time)
            return current_datapoint

    # time input will be used to interpolate an altitude value
    for key in json_data.keys():
        current_datapoint[key]=json_data[key][current_time_index]
    current_datapoint["runtime"]=int(time)
    altitude_a = json_data["altitude"][current_time_index - 1]
    altitude_b = json_data["altitude"][current_time_index]
    data_time_diff = json_data["runtime"][current_time_index] - json_data["runtime"][current_time_index - 1]
    real_time_diff = time - (json_data["runtime"][current_time_index - 1] - data_start_time)
    altitude_slope = (altitude_b - altitude_a) / data_time_diff
    altitude_result = altitude_a + altitude_slope * real_time_diff
    current_datapoint["altitude"] = altitude_result

    return current_datapoint

def get_current_sim_data(time, launch_time=5, time_to_apogee=14.30, deploy_chute_time=32.00, flight_time=55.00, apogee=1308.0, launch_altitude=20, randomness=0, engine_details=default_engine, calc_func=use_data, rocket_details=None):
    if not calc_func in [calc_altitude_sim, calc_altitude_algerbraic, use_data]:
        print("Unknown calc function. This may not be stable")
    if calc_func is calc_altitude_algerbraic:
        return calc_func(time, launch_time, time_to_apogee, deploy_chute_time, flight_time, apogee, launch_altitude, randomness)
    elif calc_func is calc_altitude_sim:
        if sim_data is None:
            reset_sim(launch_altitude, engine_details)
        return calc_func(time, launch_time, engine_details)
    elif calc_func is use_data:
        return calc_func(time)


bbp = 600
use_datafile = False 
wrapper = True
def main():
    altitude = 0
    sin, sout = create_serial_devices(gen_sout=False)
    print("sin: %s, sout: %s" % (str(sin), str(sout)))
    with os.fdopen(sin, "wb") as win:
        while 1:
            runtime = (time.time() * 1000) - start_time
            data = get_current_sim_data(runtime, calc_func=use_data if use_datafile else calc_altitude_sim)
            print(json_pretty(data, indent=4))
            if use_datafile:
                packet=DataPacket(json=data)
            else:
                packet=create_packet(
                    runtime=int(runtime),
                    accel=(int(1000 * (min(20,data["acceleration"]) if data["acceleration"] > 0 else max(-20,data["acceleration"]))),0,0),
                    alt=data["altitude"],
                    temperature=temperature
                )
            print(packet)
            if packet is None:
                continue
            #os.write(sin, packet)
            if wrapper:
                wrapped_header=WrappedHeader(json={'salt':12345, 'type':0, 'size':ctypes.sizeof(DataPacket)})
                num_out = win.write(wrapped_header)
            num_out += win.write(packet)
            print("wrote %d bytes" % (num_out))
            #for b in packet: win.write(chr(int(b)))
            win.flush()
            if bbp is None:
                time.sleep(.1)
            else:
                sleep_time = (1.024 / (bbp/num_out))
                time.sleep(sleep_time)

if __name__ == "__main__":
    main()

