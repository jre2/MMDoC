from time import sleep
import subprocess
import urllib2

PATH_SRC = 'tmp.py'
URL_SRC = 'https://raw.github.com/jre2/MMDoC/master/bot.py'
ACCOUNT_PASSWORD = open( 'password.secret' ).read()

def getSrc():
    print 'Fetching src'
    r = urllib2.urlopen( URL_SRC )
    return r.read()

def saveSrc( src ):
    print 'Saving src'
    f = open( PATH_SRC, 'w' )
    f.write( src )
    f.close()

def restart( src, p=None ):
    print 'Restarting'
    if p:
        p.kill()
    sleep( 5 )
    saveSrc( src )
    sleep( 2 )
    return subprocess.Popen( ['python', PATH_SRC, ACCOUNT_PASSWORD] )

def main():
    # start current version
    cur = getSrc()
    p = restart( cur )

    # poll for new version
    while 1:
        try:
            sleep( 30 )
            print '...polling...'
            new = getSrc()
            if new != cur:
                p = restart( new, p )
                cur = new
                print 'Updated with new code'
        except Exception, e:
            print 'Updater failed with', e

main()
