# Windows Streamdeck Media Updater

While switching between multiple PCs using a KVM switch, I wanted to be able to quickly control media playback on my personal PC.  The [Elgato Stream Deck+](https://www.elgato.com/us/en/p/stream-deck-plus) seemed like a viable candidate.  And it was... mostly.  

Out-of-the box, the Stream Deck provided standard media controls (play/stop/next/previous) and volume control.  However, there was no available solution to display what was currently playing, or to "favorite" a cool song that would come on.  

This is a quick-and-dirty solution to that problem.

### Features:
- Displays currently playing media in the touch strip (middle display) of the Stream Deck+
- Displays album art in one of the LCD keys.
- Designates one of the LCD keys as a "Favorite" button.  
    - This simply records the currently playing media to a file, for later retrieval.

### Usage:
- Ensure you have the necessary dependencies installed.  See the [comments](./getMediaInfo.py)
- Launch the script [start-media-update.bat](start-media-update.bat), either manually, or during Windows start up.  
