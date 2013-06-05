################################################################################
## Imports
################################################################################

import win32com
import win32com.client
import os
import sys
import threading
import random
from math import sqrt
from time import sleep, time
from ctypes import windll
import smtplib
from email.MIMEText import MIMEText
dc = windll.user32.GetDC( 0 )

################################################################################
## Reporting
################################################################################

GMAIL_PASSWORD = 'Ru$h2112'
GMAIL_USERNAME = 'overture2112'
GMAIL_SUBJECT = 'MMDoC Bot'

def sendmail( body ): # :: Str -> IO ()
    addr = GMAIL_USERNAME + '@gmail.com'
    recip = addr
    to = addr
    msg = MIMEText( body )
    msg['Subject'] = GMAIL_SUBJECT
    msg['To'] = recip

    s = smtplib.SMTP( 'smtp.gmail.com:587' )
    s.starttls()
    s.login( GMAIL_USERNAME, GMAIL_PASSWORD )
    s.sendmail( recip, to, msg.as_string() )
    s.quit()

################################################################################
## Screen manipulation
################################################################################

def moveMouse( x, y, fast=False ): # :: Loc -> Maybe Bool -> IO ()
    windll.user32.SetCursorPos( x, y )
    if not fast:
        sleep(0.1)

def lclick( x, y ): # :: Loc -> IO ()
    moveMouse( x, y )
    windll.user32.mouse_event( 2, 0, 0, 0, 0 ) # left down
    sleep(0.1)
    windll.user32.mouse_event( 4, 0, 0, 0, 0 ) # left up
    sleep(0.1)

def rclick( x, y ): # :: Loc -> IO ()
    moveMouse( x, y )
    windll.user32.mouse_event( 8, 0, 0, 0, 0 ) # right down
    sleep(0.2)
    windll.user32.mouse_event( 16, 0, 0, 0, 0 ) # right up
    sleep(0.2)

click = lclick

def getPixel( x, y ): # :: Loc -> Color
    c = int( windll.gdi32.GetPixel( dc, x, y ) )
    return ( c & 0xff, ( c >> 8 ) & 0xff, ( c >> 16 ) & 0xff )

def getPixelM( x, y, fast=False ): # :: Loc -> Maybe Bool -> IO Color
    moveMouse( x, y, fast )
    return getPixel( x, y )

def followLine( x, y ): # :: Loc -> IO ()
    '''Trace a verticle line until color changes [not sure in what way]'''
    def _followLine( x, y, rate ): # accurate within rate
        while abs( getPixelM( x, y, fast=True )[2] - 255 ) < 20:
            y += rate
            moveMouse( x, y, fast=True )
        return (x,y)

    x,y = _followLine( x, y, rate=10 )
    x,y = _followLine( x, y-11, rate=1 )
    return (x,y)

################################################################################
## Bot Logic Utils
################################################################################

LAST_GAME_END_TIME = time()

def dist3( a, b ):
    return sqrt( ( a[0]-b[0] )**2 + ( a[1]-b[1] )**2 + ( a[2]-b[2] )**2 )

def nearlyColor( loc, color, tolerance ):
    return dist3( getPixel( *loc ), color ) <= tolerance

def delay( sec ):
    '''Wait for roughly given amount of time'''
    t = min( 5, max( sec, random.gauss( sec, sec*0.5 ) ) )
    sleep( t )

################################################################################
## Game Specific Logic
################################################################################

COLOR_ACCEPT_HAND = (34, 61, 67)
COLOR_TURN_END = (198, 84, 16)
COLOR_TURN_WAIT = (95, 95, 95)
COLOR_CHOOSE_POSITION_CLOSE = (19, 8, 8)
COLOR_LEAVE_BUTTON = (28, 54, 61)
COLOR_HAND_POPUP_CONFIRM = (26, 47, 51)

LOC_BOARD_SPACES_CENTERS = [ (779,279), (927, 298),
                             (772,433), (923,432),
                             (767,576), (921,580),
                             (754,739), (927,736) ]
LOC_HAND_POPUP_CONFIRM = ( 647, 281 )
LOC_HERO_UNUSED_CONFIRM = ( 642, 340 )
LOC_CHOOSE_POSITION_CLOSE = (672, 198)
LOC_HAND_SCAN_START = ( 10, 900 )
LOC_HAND_SCAN_END = ( 700, 900 )
LOC_ACCEPT_HAND = ( 400, 538 )
LOC_PLAY_BUTTON = ( 948, 190 )
LOC_TURN_END = LOC_TURN_COLOR = ( 921, 73 )
LOC_QUEUE_BUTTON = ( 1361, 759 )
LOC_LEAVE_BUTTON = ( 1488, 721 )
LOC_DAILY_REWARDS_EXTEND_BUTTON = ( 1188, 834 )
LOC_HERO = ( 504, 479 )
LOC_HERO_OPTION1 = ( 500, 280 )
HERO_OPTION_DELTA = 60
LOC_ENEMY_HERO = ( 1658, 488 )
LOC_ENEMY_BACK_LANE_X = 1393
LOC_ENEMY_FRONT_LANE_X = 1220
LOC_LAUNCER_PLAY_BTN = ( 1254, 771 )
LOC_LOGIN_FOCUS = ( 994, 106 )
LOC_LOGIN_PASSWORD = ( 876, 571 )
LOC_LOGIN_BTN = ( 948, 691 )

def queueForGame():
    '''Must be out of game at home screen'''
    print 'Queueing'
    click( *LOC_PLAY_BUTTON )
    delay( 8 )
    click( *LOC_QUEUE_BUTTON )
    delay( 4 )

def acceptRewards():
    print 'Accepting rewards'
    click( *LOC_LEAVE_BUTTON )
    delay( 5 )
    click( *LOC_DAILY_REWARDS_EXTEND_BUTTON )
    delay( 8 )

def waitForGame():
    '''Wait for mulligan buttons to appear'''
    t = time()
    while 1:
        if ( time()-t ) > 60*5: raise RuntimeError( 'Spent too long waiting for game. Probably in bad state' )
        if getPixel( *LOC_ACCEPT_HAND ) == COLOR_ACCEPT_HAND:
            return
        delay( 3 )

def determineWhoseTurn():
    '''Returns None if neither were identified. Incorrect if game is over'''
    # button at top is grey and says 'Wait' while enemy turn; turns red and says 'End Turn' when your turn
    turnButtonColor = getPixel( *LOC_TURN_COLOR )
    if turnButtonColor == COLOR_TURN_END:   return 'our turn'
    if turnButtonColor == COLOR_TURN_WAIT:  return 'enemy turn'
    return 'game over'

def waitUntilTurnOrGameOver():
    print 'Waiting until my turn'
    t = time()
    while 1:
        if ( time()-t ) > 60*5: raise RuntimeError( 'Spent too long waiting for it to be our turn. Probably in bad state' )
        if getPixel( *LOC_LEAVE_BUTTON ) == COLOR_LEAVE_BUTTON: return 'game over'
        r = determineWhoseTurn()
        if r == 'enemy turn': sleep( 3 )
        else: return r

def endTurn():
    '''Try ending turn. If hero wasn't used somehow, decline msg and try again'''
    print 'Ending turn'
    # try end turn
    click( *LOC_TURN_END );             sleep( 1 )
    # if hero wasn't used, decline msg and try again
    click( *LOC_HERO_UNUSED_CONFIRM );  sleep( 1 )
    click( *LOC_TURN_END );             sleep( 1 )
    sleep( 3 )

def acceptHand():
    print 'Accepting hand'
    click( *LOC_ACCEPT_HAND )
    sleep( 3 )

def useHero( option ):
    '''Choose hero then activate option (1 indexed)'''
    print 'Using hero'
    click( *LOC_HERO )
    delay( 4 )
    x, y = LOC_HERO_OPTION1
    click( x, y + HERO_OPTION_DELTA * ( option - 1 ) )
    delay( 4 )

def tryCardsInHand():
    '''Click cards left->right until placement msg appears, then try placing in every spot in random order'''
    print 'Trying all cards in hand'
    x0, y = LOC_HAND_SCAN_START
    x1, _ = LOC_HAND_SCAN_END

    for x in xrange( x1, x0, -50 ):
        click( x, y )

        # Spells/Fortunes with popup confirmations
        sleep( 0.2 )
        if nearlyColor( LOC_HAND_POPUP_CONFIRM, COLOR_HAND_POPUP_CONFIRM, 20 ):
            print 'Accepting popup'
            click( *LOC_HAND_POPUP_CONFIRM )
            continue

        # Creatures
        locs = LOC_BOARD_SPACES_CENTERS[:]
        random.shuffle( locs )

        while nearlyColor( LOC_CHOOSE_POSITION_CLOSE, COLOR_CHOOSE_POSITION_CLOSE, 10 ): # can place
            if not locs:
                print '    Unable to play card @ %d %d' % ( x, y )
                click( *LOC_CHOOSE_POSITION_CLOSE ) # cancel card activation
                break
            loc = locs.pop()
            click( *loc )
            print '    placing %s @ %s' % ( (x,y), loc )

def attackWithAllCreatures():
    '''Click own board positions in random order to activate any available creatures,
    then click opposing enemy positions in set order'''
    print 'Attacking'
    boardLocs = LOC_BOARD_SPACES_CENTERS[:]
    random.shuffle( boardLocs )

    for x,y in boardLocs:
        # try activating creature
        click( x,y )
        # try clicking enemy hero, back lane, front lane
        for dest in [ LOC_ENEMY_HERO, ( LOC_ENEMY_BACK_LANE_X, y ), ( LOC_ENEMY_FRONT_LANE_X, y ) ]:
            click( *dest )

def mainloop( turn=0 ):
    timing, t0 = { 'me':{}, 'enemy':{} }, time()

    # Get in game
    if turn == 0:
        queueForGame()
        waitForGame()
    else:
        print 'Resuming game'
    timing['queue'] = time() - t0

    # Play game
    acceptHand()
    while 1:
        # Sanity check
        if ( time()-t0 ) > 60*25 or turn > 20: raise RuntimeError( 'Game has gone on exceptionally long. Probably in bad state' )

        # Wait for turn
        t_started_waiting = time()
        state = waitUntilTurnOrGameOver()
        if state == 'game over': break
            # minor bookeeping
        t_started_our_turn = time()
        timing['enemy'][ turn ] = t_started_our_turn - t_started_waiting

        # Play turn

            # Use hero
        delay( 1 )
        if turn <  3:   useHero( 1 ) # 4/2/2. +3 might > +1 magic > draw
        if turn == 3:   useHero( 2 )
        if turn >  3:   useHero( 4 )
        useHero( 1 )    # just in case others failed or we miscounted

        tryCardsInHand()
        attackWithAllCreatures()
        endTurn()

        # book keeping
        timing['me'][ turn ] = time() - t_started_our_turn
        turn += 1
    acceptRewards()

    # Reporting
    timing['total'] = time() - t0
    global LAST_GAME_END_TIME
    LAST_GAME_END_TIME = time()
    sendmail( 'Timing: %s' % timing )

class Bot( threading.Thread ):
    def __init__( self ):
        threading.Thread.__init__( self )
        self.daemon  = True

    def run( self ):
        print 'Starting bot'
        import pythoncom
        pythoncom.CoInitialize()

        needStart = True
        turn = 0
        if len( sys.argv ) > 1:
            needStart = False
            turn = int( sys.argv[1] )

        if needStart:
            restart()

        while 1:
            try:
                mainloop( turn )
                turn = 0
            except RuntimeError:
                restart()

def restart():
    sendmail( 'Restarting' )

    os.system('taskkill /F /IM Game.exe')
    #os.system('taskkill /F /IM Launcher.exe')
    sleep( 5 )
    #os.system(r'C:\Users\Joseph\AppData\Roaming\Ubisoft\MMDoC-PDCLive\Launcher.exe')
    #sleep( 10 )

    # start game via launcher. sometimes takes 2 clicks but usually this just messes up focus
    click( *LOC_LAUNCER_PLAY_BTN ); sleep( 1 )
    click( *LOC_LAUNCER_PLAY_BTN ); sleep( 1 )
    sleep( 15 ) # wait for login screen

    # logic
        # fix focus
    click( *LOC_LOGIN_FOCUS ); sleep( 1 )
        # select password box
    click( *LOC_LOGIN_PASSWORD ); sleep( 1 )
    click( *LOC_LOGIN_PASSWORD ); sleep( 1 ) #NOTE: probably not needed. verify this
        # type password
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('cfavader17')
        # click login
    click( *LOC_LOGIN_BTN ); sleep( 1 )
        # wait for slow load
    sleep( 20 )

    #TODO: verify at home screen, otherwise try again

def main():
    b = Bot()
    b.start()
    while 1:
        #TODO: check LAST_GAME_END_TIME and do full restart (game+bot thread)
        if windll.user32.GetAsyncKeyState( 0x76 ): # F7 key is pressed
            print 'KILL SWITCH caught'
            return
        sleep(1)

main()

'''Dev notes:
# Hang history
* Game ended during attack phase, causing it to miss leave game button and kept thinking it was still it's turn forever
* It thought game ended when it hadn't, then queued and waited forever

# Issues
* while trying to activate cards it accidently moves cards on the board sometimes
* Waits too long for certain things

# AI
* when few creatures on board, it'd be handy to toss out many small rather than 1 big
* sometimes wastes resources on spells that would be better spent getting creatures out
'''
