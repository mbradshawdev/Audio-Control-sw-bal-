# Audio-Control-sw-bal-
AHK script for switching audio devices, adjusting multi-channel balance and some other things in Win7/10.


Commands Included so far:  
'sw'itch: To select a device to set as the default device(ie switch to it), some programs and settings (exlusive mode etc) that I don't understand don't respect this and just go quiet or freeze. Relaunching the program solves this, restarting the PC has never been required. MPC-HC (my main squeeze) switches flawlessly.  
'bal'   : To enter in a set of values corresponding to the speaker channels, this pretty soon evolved into just hard-coding pre-defined settings such as bal51 for adding some volume to the rear speakers, keeping the always-too-loud sub low etc.  
  bal51, annoy, stereo, manual, default are all recognised commands for this ..useage of the script  
'shift' : Adds or subtracts a specified amount from a specified channel (or two for front and rear)  
  up, down / front, rear, left, right, back left, back right, center, sub are included ('Shift front up 10' or 'shift rear down 23')  
'set'   : As above but shifts a channel or pair of channels to a specific value ('set center 0' or 'set sub 99' etc)  
'fill'  : Toggles the enhancement 'speaker fill' programs don't seem to like changes to Enhancements, even MPC needs a currently playing track to be stopped and restarted  
'loud'  : As above but for another windows audio enhancement "Loudness Equalization" suffer the quiet voices and loud effects of movies no more!   


Posting here to demonstrate my improving coding ability (pls give me a job) and to keep active on here.  
And if anyone actually reads this: Any feedback/improvements/suggestions/employment welcome! Thanks!  
