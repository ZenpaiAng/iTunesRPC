import wmi
import presence
import win32com.client
import time
import datetime
import os
import win10toast

f = open('.pid', 'w')
f.write(str(os.getpid()))
f.close()

toaster = win10toast.ToastNotifier()

processes = []

for process in wmi.WMI().Win32_Process():
    processes.append(process.Name)

client_id = '878589532398846023'
RPC = presence.Presence(client_id, pipe=0)
try:
   RPC.connect()
except:
    toaster.show_toast("iTunesRPC", "Error: Discord Not Found.", icon_path="icon.ico", duration=3)
    time.sleep(3)
    os.sys.exit()

def get_sec(time_str):
    h, m, s = time_str.split(':')
    return int(h) * 3600 + int(m) * 60 + int(s)

oldsong = ''
wasPaused = False

if 'iTunes.exe' in processes:
    itunes = win32com.client.Dispatch('iTunes.Application')
    toaster.show_toast("iTunesRPC", "Successfully Started!", icon_path="icon.ico", duration=3)

    while True:
        if itunes.currentTrack == None:
            RPC.update(details="Not Playing", large_image="icon")

        else:
            if itunes.playerState == 0:
                wasPaused = True
                song = itunes.currentTrack.name
                artist = itunes.currentTrack.artist
    
                RPC.update(details=f"Paused", state=f"{song} by {artist}", large_image="icon")
    
            else:
                if wasPaused == True:
                    wasPaused = False
                    song = itunes.currentTrack.name
                    artist = itunes.currentTrack.artist
                    album = itunes.currentTrack.album
                    current_time = int(time.time())
                    songtime = get_sec('0:' + itunes.currentTrack.time)
                    current_pos = int(itunes.playerPosition)
                    end_time = current_time + (songtime - current_pos)
    
                    RPC.update(details=f"{song} by {artist}", state=f'from {album}', large_image="icon", start=current_time, end=end_time)
    
                else:    
                    song = itunes.currentTrack.name
                    artist = itunes.currentTrack.artist
                    album = itunes.currentTrack.album
                    timeElapsed = datetime.timedelta(seconds=itunes.playerPosition)
                    timeFinish = get_sec('0:' + itunes.currentTrack.time)
                    startTime = int(time.time()) - int(itunes.playerPosition)
            
                    if oldsong == song:
                        f = open('.end', 'r')
                        endTime = int(f.read())
                        f.close()
                    else:
                        oldsong = song
                        f = open('.end', 'w')
                        endTime = startTime + timeFinish
                        f.write(str(endTime))
                        f.close()
    
                    RPC.update(details=f"{song} by {artist}", state=f'from {album}', large_image="icon", start=startTime, end=endTime)
    
        time.sleep(15)
