MMDoC
=====

MMDoC bot and utils

#### Setup:

* Create password.secret text file with account password
* Set game to fullscreen windowed mode
* Store username for login

#### Usage:
1. Start game launcher (but not the game itself) and don't move window from default position
2. Run `python driver.py` from base directory via a command prompt

Hit F7 to kill bot process (driver.py will remain alive)

Dev Notes
=========
#### Hang history
* Game ended during attack phase, causing it to miss leave game button and kept thinking it was still it's turn forever
* It thought game ended when it hadn't, then queued and waited forever

#### Issues
* while trying to activate cards it accidently moves cards on the board sometimes
* Waits too long for certain things

#### AI
* when few creatures on board, it'd be handy to toss out many small rather than 1 big
* sometimes wastes resources on spells that would be better spent getting creatures out
