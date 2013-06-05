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
dc = windll.user32.GetDC( 0 ) # TODO: replace with exact dc

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

LAST_GAME_END_TIME = time()
COLOR_ACCEPT_HAND = (34, 61, 67)
COLOR_TURN_END = (198, 84, 16)
COLOR_TURN_WAIT = (95, 95, 95)
COLOR_CHOOSE_POSITION_CLOSE = (19, 8, 8)
COLOR_LEAVE_BUTTON = (28, 54, 61)

LOC_BOARD_SPACES_CENTERS = [ (779,279), (927, 298),
                             (772,433), (923,432),
                             (767,576), (921,580),
                             (754,739), (927,736) ]
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
    click( *LOC_TURN_END ); sleep( 1 )
    # if hero wasn't used, decline msg and try again
    click( 642, 340 );      sleep( 1 )
    click( *LOC_TURN_END ); sleep( 1 )
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
        if nearlyColor( (647, 281), (26, 47, 51), 20 ): #FIXME: const
            print 'Accepting popup'
            click( 647, 281 ) # accept if there's a popup
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
        for dest in [ (1658,488), (1393,y), (1220,y) ]:
            click( *dest )

def mainloop( turn=0 ):
    timing, t0 = {}, time()

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
        if ( time()-t0 ) > 60*25 or turn > 20: raise RuntimeError( 'Game has gone on exceptionally long. Probably in bad state' )
        # Wait for turn
        state = waitUntilTurnOrGameOver()
        if state == 'game over': break
        t = time()
        
        # Play turn

            # Use hero
        delay( 1 )
        if turn <  3:   useHero( 1 ) # 4/2/2 then draw
        if turn == 3:   useHero( 2 )
        if turn >  3:   useHero( 4 )
        useHero( 1 )    # just in case others failed or we miscounted
        
        tryCardsInHand()
        attackWithAllCreatures()
        endTurn()
        
        # book keeping
        timing[ turn ] = time() - t
        turn += 1
        #TODO: if time goes past X min, bail and restart
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
    
    click( 1254, 771 ); sleep( 1 ) # play btn
    click( 1254, 771 ); sleep( 1 ) # play btn
    
    sleep( 15 ) # wait for login screen

    click( 994, 106 ) # get focus for mmdoc
    click( 876, 571 ); sleep( 1 ) # password box
    click( 876, 571 ); sleep( 1 ) # password box
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('cfavader17')
    click( 948, 691 ); sleep( 1 ) # login btn
    sleep( 20 ) # wait for login
    #FIXME: verify at home screen, otherwise try again

def main():
    b = Bot()
    b.start()
    while 1:
        if windll.user32.GetAsyncKeyState( 0x76 ): # F7 key is pressed
            print 'KILL SWITCH caught'
            return
        sleep(1)

main()
