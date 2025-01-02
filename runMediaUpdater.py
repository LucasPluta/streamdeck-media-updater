#!/usr/bin/env python3

import threading
import io
import time
import asyncio
import traceback
import os

from PIL import Image, ImageDraw, ImageFont
from StreamDeck.DeviceManager import DeviceManager
from StreamDeck.Transport.Transport import TransportError
from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as MediaManager
from winsdk.windows.storage.streams import DataReader, Buffer, InputStreamOptions

'''This script will display the currently playing media on the Stream Deck Touchscreen and the album art on Key 6'''

# Required pip packages
#   winsdk - Installing this pip package requires Windows Developer Tools provided by Visual Studio
#   streamdeck

# Also requires this: https://python-elgato-streamdeck.readthedocs.io/en/stable/pages/backend_libusb_hidapi.html#windows

ALBUM_ART_KEY_NUMBER = 6
REFRESH_BUTTON_KEY_NUMBER = 5

async def get_media_info():
    '''Polls the WinRT API for the currently playing media info'''
    sessions = await MediaManager.request_async()

    current_session = sessions.get_current_session()
    if current_session:  # there needs to be a media session running
        #if current_session.source_app_user_model_id == TARGET_ID:
        info = await current_session.try_get_media_properties_async()

        # song_attr[0] != '_' ignores system attributes
        info_dict = {song_attr: info.__getattribute__(song_attr) for song_attr in dir(info) if song_attr[0] != '_'}

        # converts winrt vector to list
        info_dict['genres'] = list(info_dict['genres'])

        thumb_stream_ref = info_dict['thumbnail']

        try:
            # Copy the thumbnail image from the stream reference that is provided by the API
            thumb_read_buffer = Buffer(5000000)
            readable_stream = await thumb_stream_ref.open_read_async()
            await readable_stream.read_async(thumb_read_buffer, thumb_read_buffer.capacity, InputStreamOptions.READ_AHEAD)
            img = Image.new('RGB', (120, 120), color='black')
            released_icon_from_buffer = Image.open(io.BytesIO(bytearray(thumb_read_buffer))).resize((120, 120))
            img.paste(released_icon_from_buffer, (0, 0))
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_released_bytes = img_byte_arr.getvalue()
            info_dict['thumbnailBytes'] = img_released_bytes
            del readable_stream
            del thumb_read_buffer
            del img
        except:
            #traceback.print_exc()
            info_dict['thumbnailBytes'] = None

        # Deallocate all objects to avoid memory leaks
        del sessions
        del current_session
        del info
        del thumb_stream_ref
        
        return info_dict

    info_dict = {}
    info_dict['title'] = "---"
    info_dict['thumbnailBytes'] = "No Image"

    # Deallocate all objects to avoid memory leaks
    del sessions
    del current_session
    
    return info_dict

def get_current_media_info():
    '''Gets the current media info'''
    try:
        current_media_info = asyncio.run(get_media_info())
        return current_media_info
    except:
        traceback.print_exc()
        current_media_info = {}
        current_media_info['title'] = ""
        return current_media_info

def updateCurrentlyPlaying(deck, current_media_info):  
    '''Updates the currently playing media info on the Stream Deck Touchscreen'''
    title = current_media_info['title']
    if title == None or title == "":
        print("Not updating... title is empty")
        print(current_media_info)
        return

    with deck:
        try:
            img = Image.new('RGB', (600, 100), 'black')
            
            # Strip any non-ASCII characters
            title = ''.join([i if ord(i) < 128 else '' for i in title])
            # Trim the whitespace from beginning and end
            title = title.strip()
            # split long titles into multiple lines, perferably at a space
            
            if len(title) > 65:
                title = title[:65] + '\n' + title[65:]
            text = title
            
            # Check if current_media_info has an "artist"
            if 'artist' in current_media_info:
                text += "\n" + current_media_info['artist']
            # Check if current_media_info has an "album"
            if 'album_title' in current_media_info:
                text += "\n" + current_media_info['album_title']

            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("C:\\WINDOWS\\FONTS\\ARIALBD.ttf", 18)
            draw.text((img.width / 1.6, 15), font=font, text=text,  anchor="ms", fill="white")

            img_bytes = io.BytesIO()
            img.save(img_bytes, format='JPEG')
            touchscreen_image_bytes = img_bytes.getvalue()

            deck.set_touchscreen_image(touchscreen_image_bytes, 200,0,600,100) 
            print("Setting deck to: " + title) 

        except:
            print ("Error updating currently playing media")
            traceback.print_exc()

def updateAlbumArt(deck, media_info):
    '''Updates the album art on the Stream Deck Key 6'''
    if media_info == None or "thumbnailBytes" not in media_info:
        return None

    image = media_info['thumbnailBytes']
    if image == None:
        return

    if image == "No Image":
        blankAlbumArt(deck)
        return

    with deck:
        try:
            deck.set_key_image(ALBUM_ART_KEY_NUMBER, image)
            print("Album art updated")
        except Exception: 
            traceback.print_exc()
            img = Image.new('RGB', (120, 120), color='black')
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_released_bytes = img_byte_arr.getvalue()
            deck.set_key_image(ALBUM_ART_KEY_NUMBER, img_released_bytes)
            print("Album art not updated")
        
def hashAlbumArt(media_info):
    '''Compares two album arts to see if they are the same'''
    if media_info == None or "thumbnailBytes" not in media_info:
        return None

    image = media_info['thumbnailBytes']
    if image == None:
        return

    try:
        return hash(media_info['thumbnailBytes'])
    except:
        return None

def key_change_callback(deck, key, key_state):
    '''Callback for when a key is pressed or released'''
    if key == REFRESH_BUTTON_KEY_NUMBER and key_state:
        current_media_info = get_current_media_info()
        updateCurrentlyPlaying(deck, current_media_info)
        blankAlbumArt(deck)
        updateAlbumArt(deck, current_media_info)

    if key == ALBUM_ART_KEY_NUMBER and key_state:
        current_media_info = get_current_media_info()
        title = current_media_info['title']
        artist = current_media_info['artist']
        album = current_media_info['album_title']
        recordFavorite(title, artist, album)
        

def recordFavorite(title, artist, album):
    favoriteLine = "\"" + title + "\",\"" + artist + "\",\"" + album + "\"\n"
    print("FAVORITE: " + favoriteLine)
    # Check if a favorites.csv file exists, if not create one and write the header line
    if not os.path.isfile("favorites.csv"):
        file = open("favorites.csv", "w")
        file.write("Title,Artist,Album\n")
        file.close()

    # Check if the last line of the csv is the favorite
    file = open("favorites.csv", "r")
    lines = file.readlines()
    file.close()
    if len(lines) > 1:
        lastLine = lines[-1]
        if lastLine == favoriteLine:
            return

    file = open("favorites.csv", "a")
    file.write(favoriteLine)
    file.close()

def blankAlbumArt(deck):
    with deck:
        img = Image.new('RGB', (120, 120), color='black')
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG')
        img_released_bytes = img_byte_arr.getvalue()
        deck.set_key_image(ALBUM_ART_KEY_NUMBER, img_released_bytes)

def runUpdaterTask(deck):
    previousTitle = ""
    previousMediaHash = None
    while True:
        time.sleep(0.500)
        current_media_info = get_current_media_info()
        currentTitle = current_media_info["title"]
        if previousTitle != currentTitle and currentTitle != None and currentTitle != "":
            updateCurrentlyPlaying(deck, current_media_info) 
            blankAlbumArt(deck)
            previousMediaHash = None     
            previousTitle = currentTitle
        
        currentMediaHash = hashAlbumArt(current_media_info)
        if previousMediaHash != currentMediaHash and currentMediaHash != None:
             updateAlbumArt(deck, current_media_info)
             previousMediaHash = currentMediaHash
                
if __name__ == "__main__":
    while True:
        streamdecks = DeviceManager().enumerate()

        print("Found {} Stream Deck(s).\n".format(len(streamdecks)))

        for index, deck in enumerate(streamdecks):
            # This example only works with devices that have screens.

            if deck.DECK_TYPE != 'Stream Deck +':
                print(deck.DECK_TYPE)
                print("Sorry, this example only works with Stream Deck +")
                continue

            deck.open()
            #deck.reset()

            deck.set_key_callback(key_change_callback)

            print("Opened '{}' device (serial number: '{}')".format(deck.deck_type(), deck.get_serial_number()))

            try:
                runUpdaterTask(deck)
            except (TransportError, RuntimeError):
                print("Error running updater task")
            
            for t in threading.enumerate():
                try:
                    t.join()
                except (TransportError, RuntimeError):
                    pass
        print("No stream Deck Detected...")
        time.sleep(1)
