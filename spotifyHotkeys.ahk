#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#Include %A_ScriptDir%
#Include Spotify.ahk
SetTimer,TOOLTIP,5000
SetTimer,TOOLTIP,Off

; /* ============== */ 
; /*     Spotify    */
; /* ============== */
; Download https://github.com/CloakerSmoker/Spotify.ahk
; Add shortcut to shell:Startup
; Add lib folder to shell:Startup
; In Windows settings, set .mp3 association to Spotify Windows App (Spotify.exe) by right clicking an mp3 file, selecting Open With -> Choose another app, Look for another app on this PC, find Spotify.exe. If you don't do this, Windows might open Windows Media Player or another media app whenever you press a media key.
; If you're using the Windows app, run this in PowerShell (as admin) to get exact location of Spotify.exe: "$((ls "C:\Program Files\WindowsApps\" | Where-Object { $_ -like "*spotify*"})[0].FullName)\Spotify.exe"
; If Spotify updated, the folder changed and Windows Media Player or another app may open when you press your media keys. Redo the above two steps.
; If you get an error message attempting to set Spotify to open mp3s, open Task Manager with CTRL + Shift + Esc, click File -> Open New Task, check "Create this task with administrative privileges", type Explorer, and press enter. Try it from there. Then close Explorer and see if the media keys work.
; https://community.spotify.com/t5/Desktop-Windows/Keyboard-media-launch-button/td-p/1736288

; Create new Spotify object
spotifyObject := new Spotify

; Activate Spotify whenever modifier is configured
*Launch_Media::
  Run Spotify
return

; Shift + Media_Play_Pause adds the song to a playlist via menu to display all playlist options
+Media_Play_Pause:: 
  playlistMenu()
Return

; Media_Play_Pause toggles playback and shows Now Playing information in a tool tip
*Media_Play_Pause:: 
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  Send, {Media_Play_Pause}
  showSongInfo(CurrentPlayback)
Return

; CTRL + Media_Play_Pause just gives Now Playing information in a tool tip
^Media_Play_Pause:: 
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  showSongInfo(CurrentPlayback)
Return

; Shift + CTRL + Media_Play_Pause gives Now Playing information in a tool tip and pastes to clipboard
^+Media_Play_Pause:: 
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  clipboard:= showSongInfo(CurrentPlayback)
Return

; Media_Next shows song info when going to next track
*Media_Next::
  Send,{Media_Next}
  Sleep, 1000
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  showSongInfo(CurrentPlayback)
return
; CTRL + Next seeks forward ~10 seconds
^Media_Next:: 
  CurrentPlayer := spotifyObject.Player
  CurrentPlayback := CurrentPlayer.GetCurrentPlaybackInfo()
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms + 10000)
  ToolTip,% "Forward ~10 seconds"
  SetTimer,TOOLTIP,On
Return
; Shift + Next seeks forward ~30 seconds
+Media_Next:: 
  CurrentPlayer := spotifyObject.Player
  CurrentPlayback := CurrentPlayer.GetCurrentPlaybackInfo()
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms + 10000 * 3)
  ToolTip,% "Forward ~30 seconds"
  SetTimer,TOOLTIP,On
Return

; Show song info when going to previous track
*Media_Prev::
  Send,{Media_Prev}
  Sleep, 1000
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  showSongInfo(CurrentPlayback)
return
; CTRL + Prev seeks backwards ~10 seconds
^Media_Prev::
  CurrentPlayer := spotifyObject.Player
  CurrentPlayback := CurrentPlayer.GetCurrentPlaybackInfo()
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms - 10000)
  ToolTip,% "Backwards ~10 seconds"
  SetTimer,TOOLTIP,On
Return
; Shift + Prev seeks backwards ~30 seconds
+Media_Prev::
  CurrentPlayer := spotifyObject.Player
  CurrentPlayback := CurrentPlayer.GetCurrentPlaybackInfo()
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms - 10000 * 3)
  ToolTip,% "Backwards ~30 seconds"
  SetTimer,TOOLTIP,On
Return

; CTRL + Media button adds to New School and Social Media Likes playlists
^Launch_Media::
  addSongToPlaylist("id", "New School")
  Sleep, 50 
  addSongToPlaylist("id", "Social Media Likes")
  SetTimer,TOOLTIP,On 
Return

; Shift + Media button adds to Social Media Likes playlists
+Launch_Media::
  addSongToPlaylist("id", "Social Media Likes")
  SetTimer,TOOLTIP,On 
Return

; CTRL + Shift + Media button adds to Starred playlist
^+Launch_Media::
  addSongToPlaylist("id", "Starred")
Return

; Media button adds to Starred2 and Starred3 playlists
Launch_Media:: 
  ; Starred2 is my primary playlist for new songs I like
  addSongToPlaylist("id", "Starred2")
  ; Adding new tracks to Starred3 as well in case tracks added to Starred2 are reset at some point in time since the playlist is over 10k
  Sleep, 50 
  addSongToPlaylist("id", "Starred3")
Return

; Helps display Tooltip
TOOLTIP:
  ToolTip,
  SetTimer,TOOLTIP,Off
Return

; Add the song in CurrentPlayback to the playlist "name" identified by playlistID
addSongToPlaylist(playlistID, name){
  spotifyObject := new Spotify
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  ToolTip,% "Adding the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your " name " playlist"
  SetTimer,TOOLTIP,On 
  playlistObject := spotifyObject.Playlists.GetPlaylist(playlistID)
  playlistObject.AddTrack(CurrentPlayback.Track.ID)
  ToolTip,% "Added the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your " name " playlist"
  SetTimer,TOOLTIP,On 
}

; Displays information about the current song (title, artists, album, progress)
showSongInfo(CurrentPlayback)
{ 
  artistString := ""
  for index, element in CurrentPlayback.Track.artists
  {
    artistString .= ", " . CurrentPlayback.Track.artists[index].name
  }
  artistString := substr(artistString,2)
  songInfo := "Title: " CurrentPlayback.Track.Name " `nArtist: " artistString " `nAlbum: " CurrentPlayback.Track.album.name " `nProgress: " floor(CurrentPlayback.progress_ms/1000/60)":"floor(mod(CurrentPlayback.progress_ms/1000,60)) " - " floor(CurrentPlayback.Track.duration/1000/60)":"floor(mod(CurrentPlayback.Track.duration/1000,60))
  ToolTip,%songInfo%
  SetTimer,TOOLTIP,On 
return songInfo
}

; Build Spotify playlist menu
playlistMenu(){ 
	; Default option at top of menu is to display song information
  spotifyObject := new Spotify
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  boundParam := Func("showSongInfo").Bind(CurrentPlayback)
  ; Construct the rest of the menu
  Menu, convert, Add
  Menu, convert, Delete	 
  Menu, convert, Add, Add the currently playing Spotify song to a playlist..., % boundParam
  Menu, convert, Add,
  Menu, convert, Add, &Enter your playlist name here, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Enter your playlist2 name here, SPOTIFY_MENU_ACTION
  Menu, convert, Default, Add the currently playing Spotify song to a playlist...
  Menu, convert, Show
Return
}

; Return playlist ID by playlist name
getPlaylistID(playlistName){
  If (playlistName == "&Starred"){
    return "id"
  } Else If (playlistName == "&Enter your playlist name here"){
    return "enter id of your playlist here"
  } Else If (playlistName == "&Enter your playlist2 name here"){
    return "enter id of your playlist2 here"
  } 
}

; Pass selection to Spotify API to be added to playlist
SPOTIFY_MENU_ACTION(){
  playlistID := getPlaylistID(A_ThisMenuItem)
  addSongToPlaylist(playlistID, A_ThisMenuItem)
Return
}

; 
; Caveats
; 
; The script will add duplicate songs to a playlist without checking and the Spotify API won't stop me. Tried to implement logic [If NewSchool.tracks [not] contains CurrentPlayback.Track.ID] but it looks like this only checks the first 100 tracks returned from the playlist. However I noticed this puts me over the 10k song limit? Starred2 has over 10k songs adding songs using this script while adding them via the Spotify player gives me an error message. 
; 
; CTRL + Play was supposed to shuffle devices
; This API appears to be bugged because it only returns my current active playback device, not my other available Devices
; https://developer.spotify.com/console/get-users-available-devices/
; https://developer.spotify.com/documentation/web-api/guides/using-connect-web-api/#devices-not-appearing-on-device-list
; https://community.spotify.com/t5/Spotify-for-Developers/v1-me-player-devices-returns-empty-array/m-p/5224904
; https://github.com/spotify/web-api/issues/671
; ^Media_Play_Pause:: 
;   CurrentPlayer := spotifyObject.Player
;   Devices := CurrentPlayer.GetDeviceList
;   Sleep, 1
; [...]
; Return
; 
