#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#Include %A_ScriptDir%
#Include Spotify.ahk
SetTimer,TOOLTIP,3000
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

; This API appears to be bugged because it only returns my current active playback device, not my other available Devices
; https://developer.spotify.com/console/get-users-available-devices/
; https://developer.spotify.com/documentation/web-api/guides/using-connect-web-api/#devices-not-appearing-on-device-list
; https://community.spotify.com/t5/Spotify-for-Developers/v1-me-player-devices-returns-empty-array/m-p/5224904
; https://github.com/spotify/web-api/issues/671
; https://github.com/spotify/web-api/issues/671
; CTRL + Play shuffles devices
; ^Media_Play_Pause:: 
;   CurrentPlayer := spotifyObject.Player
;   Devices := CurrentPlayer.GetDeviceList
;   Sleep, 1
; [...]
; Return

; CTRL + Next seeks forward ~10 seconds
^Media_Next:: 
  CurrentPlayer := spotifyObject.Player
  CurrentPlayback := CurrentPlayer.GetCurrentPlaybackInfo()
  CurrentPlayer.SeekTime(CurrentPlayback.progress_ms + 11000)
  ToolTip,% "Forward 10 seconds"
  SetTimer,TOOLTIP,On
Return

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
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  ToolTip,% "Adding the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your New School & Social Media Likes playlists"
  SetTimer,TOOLTIP,On 
  NewSchool := spotifyObject.Playlists.GetPlaylist("playistID")
  SocialMediaLikes := spotifyObject.Playlists.GetPlaylist("playistID")
  NewSchool.AddTrack(CurrentPlayback.Track.ID)
  SocialMediaLikes.AddTrack(CurrentPlayback.Track.ID)
  ToolTip,% "Added the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your New School & Social Media Likes playlists"
  SetTimer,TOOLTIP,On 
Return

; Shift + Media button adds to Social Media Likes playlists
+Launch_Media::
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  ToolTip,% "Adding the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your Social Media Likes playlist"
  SetTimer,TOOLTIP,On 
  SocialMediaLikes := spotifyObject.Playlists.GetPlaylist("playistID")
  SocialMediaLikes.AddTrack(CurrentPlayback.Track.ID)
  ToolTip,% "Added the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your Social Media Likes playlists"
  SetTimer,TOOLTIP,On 
Return

; CTRL + Shift + Media button adds to Starred playlist
^+Launch_Media::
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo()
  ToolTip,% "Adding the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your Starred playlist"
  SetTimer,TOOLTIP,On 
  Starred := spotifyObject.Playlists.GetPlaylist("playistID")
  Starred.AddTrack(CurrentPlayback.Track.ID)
  ToolTip,% "Added the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your Starred playlist"
  SetTimer,TOOLTIP,On 
Return

; Media button adds to Starred2 and Starred3 playlists
Launch_Media:: 
  CurrentPlayback := spotifyObject.Player.GetCurrentPlaybackInfo() 
  ToolTip,% "Adding the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your Starred2 & Starred3 playlists"
  SetTimer,TOOLTIP,On
  Starred2 := spotifyObject.Playlists.GetPlaylist("playlistID") 
  Starred3 := spotifyObject.Playlists.GetPlaylist("playlistID")
  ; If NewSchool.tracks [not] contains CurrentPlayback.Track.ID ; Looks like this only checks the first 100 tracks returned from the playlist so could have duplicates. However I noticed this puts me over the 10k song limit?
  Starred2.AddTrack(CurrentPlayback.Track.ID)
  Starred3.AddTrack(CurrentPlayback.Track.ID)
  ToolTip,% "Added the song """ CurrentPlayback.Track.Name """ by " CurrentPlayback.Track.artists[1].name " to your Starred2 & Starred3 playlists"
  SetTimer,TOOLTIP,On
Return

TOOLTIP:
  ToolTip,
  SetTimer,TOOLTIP,Off
Return
