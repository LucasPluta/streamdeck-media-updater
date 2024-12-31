#!/usr/bin/env python3

import threading
import io
import time
import asyncio

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

async def read_stream_into_buffer(stream_ref, buffer):
    '''Reads the contents of a WinRT stream type into a buffer'''
    readable_stream = await stream_ref.open_read_async()
    readable_stream.read_async(buffer, buffer.capacity, InputStreamOptions.READ_AHEAD)

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

        # Deallocate all objects to avoid memory leaks
        del sessions
        del current_session
        del info

        return info_dict

    info_dict = {}
    info_dict['title'] = ""

    # Deallocate all objects to avoid memory leaks
    del sessions
    del current_session
    
    return info_dict

def get_current_media_info():
    '''Gets the current media info'''
    try:
        current_media_info = asyncio.run(get_media_info())
        print( current_media_info)
        return current_media_info
    except:
        current_media_info = {}
        current_media_info['title'] = ""
        return current_media_info

def updateCurrentlyPlaying(deck, current_media_info):  
    '''Updates the currently playing media info on the Stream Deck Touchscreen'''
    with deck:
        try:
            print("Updating currently playing media")
            img = Image.new('RGB', (600, 100), 'black')
            
            title = current_media_info['title']
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

            # Deallocate all objects to avoid memory leaks
            del img
            del draw
            del font
            del img_bytes
            del touchscreen_image_bytes
            
        except:
            print ("Error updating currently playing media")


def updateAlbumArt(deck, current_media_info):
    '''Updates the album art on the Stream Deck Key 6'''
    with deck:
        print("Updating album art")
        try:
            # Set the key image to the thumbnail of the currently playing media
            thumb_stream_ref = current_media_info['thumbnail']
            thumb_read_buffer = Buffer(5000000)
            asyncio.run(read_stream_into_buffer(thumb_stream_ref, thumb_read_buffer))
            buffer_reader = DataReader.from_buffer(thumb_read_buffer)
            byte_buffer = buffer_reader.read_bytes(thumb_read_buffer.length)
            img = Image.new('RGB', (120, 120), color='black')
            released_icon_from_buffer = Image.open(io.BytesIO(bytearray(byte_buffer))).resize((120, 120))
            img.paste(released_icon_from_buffer, (0, 0))
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_released_bytes = img_byte_arr.getvalue()
            deck.set_key_image(ALBUM_ART_KEY_NUMBER, img_released_bytes)

            # Deallocate all objects to avoid memory leaks
            del img
            del thumb_read_buffer
            del buffer_reader
            del byte_buffer
            del img_byte_arr
            del released_icon_from_buffer

        except:
            img = Image.new('RGB', (120, 120), color='black')
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG')
            img_released_bytes = img_byte_arr.getvalue()
            deck.set_key_image(ALBUM_ART_KEY_NUMBER, img_released_bytes)

            # Deallocate all objects to avoid memory leaks
            del img
            del img_byte_arr
            del img_released_bytes

def hashAlbumArt(media_info):
    '''Compares two album arts to see if they are the same'''
    try:
        t1Buf = Buffer(5000000)
        asyncio.run(read_stream_into_buffer(media_info['thumbnail'], t1Buf))
        t1Reader = DataReader.from_buffer(t1Buf)
        t1ByteBuf = t1Reader.read_bytes(t1Buf.length)
        t1ByteArr = bytes(bytearray(t1ByteBuf))

        # Deallocate all objects to avoid memory leaks
        del t1Buf
        del t1Reader

        return hash(t1ByteArr)
    except:
        return None

def key_change_callback(deck, key, key_state):
    '''Callback for when a key is pressed or released'''
    print("Key {} has been {}".format(key, "pressed" if key_state else "released"))
    if key == REFRESH_BUTTON_KEY_NUMBER and key_state:
        current_media_info = get_current_media_info()
        updateCurrentlyPlaying(deck, current_media_info)
        updateAlbumArt(deck, current_media_info)
        
        # Deallocate all objects to avoid memory leaks
        del current_media_info

def runUpdaterTask(deck):
    previous_media_info = get_current_media_info()
    previous_media_info["title"] = None
    previousMediaHash = None
    while True:
        time.sleep(0.250)
        current_media_info = get_current_media_info()
        if current_media_info["title"] != previous_media_info["title"]:
            updateCurrentlyPlaying(deck, current_media_info)      
            previous_media_info = current_media_info
        if previousMediaHash != hashAlbumArt(current_media_info):
             print(previous_media_info)
             print(current_media_info)
             updateAlbumArt(deck, current_media_info)
             previousMediaHash = hashAlbumArt(current_media_info)
        
        # Deallocate all objects to avoid memory leaks
        del current_media_info
        
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
