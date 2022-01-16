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
; In Windows settings, set .mp3 association to Spotify Windows App: C:\Program Files\WindowsApps\SpotifyAB.SpotifyMusic_1.173.517.0_x86__zpdnekdrzrea0\Spotify.exe
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
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms + 11000)
  ToolTip,% "Forward 10 seconds"
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
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms - 11000)
  ToolTip,% "Backwards 10 seconds"
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
  Menu, convert, Add
  Menu, convert, Delete	 
  Menu, convert, Add, Add the currently playing Spotify song to a playlist..., SPOTIFY_MENU_ACTION
  Menu, convert, Add,
  Menu, convert, Add, &Starred, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Starred2, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Starred3, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Social Media Likes, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &New School, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Instrumentals, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Ryan's Lofi, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Video Game Rap, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &Not Rap, SPOTIFY_MENU_ACTION
  Menu, convert, Add, &My Current Playlist, SPOTIFY_MENU_ACTION
  Menu, convert, Default, Add the currently playing Spotify song to a playlist...
  Menu, convert, Show
Return
}

; Return playlist ID by playlist name
getPlaylistID(playlistName){
  If (playlistName == "&Starred"){
    return "id"
  } Else If (playlistName == "&Starred2"){
    return "id"
  } Else If (playlistName == "&Starred3"){
    return "id"
  } Else If (playlistName == "&Social Media Likes"){
    return "id"
  } Else If (playlistName == "&New School"){
    return "id"
  } Else If (playlistName == "&Instrumentals"){
    return "id"
  } Else If (playlistName == "&Ryan's Lofi"){
    return "id"
  } Else If (playlistName == "&Video Game Rap"){
    return "id"
  } Else If (playlistName == "&Not Rap"){
    return "id"
  } Else If (playlistName == "&My Current Playlist"){
    return "id"
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
