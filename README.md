# email-stats
A python script that produces some statistics on your email usage.

![Screenshot](/email_analyze.png?raw=true "Screenshot")

You can use it with your gmail account (IMAP access needs to be enabled first). The script file contains a config section, your password etc. needs to be filled in there
(create a app-specific password for higher security - and obviously do not share it with anyone). You only need to fetch once, the headers can then be saved in a file (an option in the config).
 
The script can be run with the comman parameter "-s" or "--save", then the plot is saved to a png file. The default behavior is to display it.

Requires seaborn version >=0.6 (dev at this time).

Released under GPL.


## Related projects:
- [Gmail Meter](http://gmailmeter.com/)
- [Glowing Python blog post](http://glowingpython.blogspot.com/2012/05/analyzing-your-gmail-with-matplotlib.html)
