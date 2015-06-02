########
#
# analyzes email from (IMAP/gmail)
#
#
# some stuff was originally based on http://glowingpython.blogspot.co.at/2012/05/analyzing-your-gmail-with-matplotlib.html
#
# requires seaborn version >=0.6 (dev at this time)
#
# GPL, by A. Riss
#
########


import imaplib
import dateutil
import pickle
import sys
import os.path
import string
import email
import datetime
import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import seaborn as sns


# CONFIG

CFG_online_data_save = True   # save online data to file
CFG_offline_data = True       # use previously stored offline data
CFG_offline_data_filename = "headers_allmail.p"   # file to store server data (or load previously fetched server data - if CFG_offline_data is True)

CFG_mail_imap_server = 'imap.gmail.com'  # IMAP server
CFG_mail_account = "gmail@chucknorris.com"
CFG_mail_passwd = "thesecretpassword"

CFG_imap_folder = '"[Gmail]/All Mail"'   # imap folder to analyze
CFG_imap_download_lastdays = 365 * 5    # download emails from last x days

CFG_name_analyze = "Chuck Norris"  # to distinguish sent and received emails, probably better to analyze the headers
CFG_date_start = datetime.datetime(2013,9,1, tzinfo=dateutil.tz.tzlocal())

CFG_header_fields = ['date', 'sender_email', 'sender_name', 'receipient_email', 'receipient_name', 'receipient_cc', 'subject', 'message_id', 'message_id_intern']

CFG_figsize = (16, 8)

CFG_text_mailssent = "Mails sent"
CFG_text_mailsreceived = "Mails received"
CFG_text_weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']





def getHeaders(address,password,folder,d):
    """ retrieve the headers of the emails, starting from d days ago."""
    imaplib._MAXLINE = 200000  # some hack to increase max number of bytes
    mail = imaplib.IMAP4_SSL(CFG_mail_imap_server)
    mail.login(address,password)
    mail.select(folder, readonly=True) 
    
    # retrieving the uids
    interval = (datetime.date.today() - datetime.timedelta(d)).strftime("%d-%b-%Y")
    result, data = mail.uid('search', None, '(SENTSINCE {date})'.format(date=interval))
    
    # retrieving the headers
    result, data = mail.uid('fetch', data[0].decode().replace(' ',','), '(BODY.PEEK[HEADER.FIELDS (DATE FROM TO CC SUBJECT MESSAGE-ID)])')
    mail.close()
    mail.logout()
    
    return data

    
def parseHeaders(headers):
    """parses email headers and returns a pandas dataframe"""
    
    df = {}
    for n in CFG_header_fields: df[n] = []
    
    datesCorrected = 0
    mailsSkipped = 0
    mailsAnalyzed = 0
    for h in headers:
        if len(h)<2: continue
        h = h[1]
        
        # initialize variables
        mailstamp = datetime.datetime(datetime.MINYEAR, 1, 1)
        subject = ""
        message_id = ""
        from_name = ""
        from_email = ""
        to_name = []
        to_email = []
        to_cc = []  # list of True/False, recipient will be in the to_name and to_email fields

        # lines = email.message_from_string(h)
        lines = email.message_from_bytes(h)
        
        try:
            message_id = lines['message-id']
            subject = lines['subject']
            tos = lines.get_all('to', [])
            ccs = lines.get_all('cc', [])
            froms = lines.get('from', "")
            for n, e in email.utils.getaddresses(tos):
                to_name.append(n)
                to_email.append(e)
                to_cc.append(False)
            for n, e in email.utils.getaddresses(ccs):
                to_name.append(n)
                to_email.append(e)
                to_cc.append(True)
            from_name  = email.utils.parseaddr(froms)[0]
            from_email = email.utils.parseaddr(froms)[1]
            
            date_str = str(lines['date'])
            try:
                mailstamp = email.utils.parsedate_to_datetime(date_str)
            except TypeError:
                datesCorrected+=1
                filter(lambda x: x in string.printable, date_str)
                date_str = date_str.replace("MET DST","CEST")  # parser gets confused by MET DST
                date_str = date_str.replace("METDST","CEST")
                date_str = date_str.replace("Westeuropische Normalzeit","CET")
                date_str = date_str.replace("Mitteleuropische Zeit","CET")
                mailstamp = dateutil.parser.parse(date_str, fuzzy=True)
            mailstamp = mailstamp.astimezone(dateutil.tz.tzlocal())
        except (TypeError,ValueError):
            mailsSkipped +=1
            continue
        
        mailsAnalyzed += 1
        if len(to_name)==0:  # sometimes we do not have any visible recipients
            to_name.append("")
            to_email.append("")
            to_cc.append(False)
        for n,e,c in zip(to_name,to_email,to_cc):
            df['date'].append(mailstamp)
            df['sender_email'].append(from_email)
            df['sender_name'].append(from_name)
            df['receipient_email'].append(e)
            df['receipient_name'].append(n)
            df['receipient_cc'].append(c)
            df['subject'].append(subject)
            df['message_id'].append(message_id)
            df['message_id_intern'].append(mailsAnalyzed)  # use sequential number to make our own id
            
    print('Corrected %s timestamps.' % datesCorrected)
    print('Skipped %s mails.' % mailsSkipped)
    print('Successfully parsed %s mails.' % mailsAnalyzed)
    
    df =  pd.DataFrame(df)
    df['date'] = pd.to_datetime(df['date'])
    return df
    
    
def weekdayPlot(df, days_span):
    """plots counts of emails per weekday"""
    
    fig = plt.figure(num=None, figsize=CFG_figsize)
    gs = mpl.gridspec.GridSpec(2, 2,height_ratios=[1,3])
    ax = []
    ax.append(plt.subplot(gs[0]))
    ax.append(plt.subplot(gs[1], sharey=ax[0], sharex=ax[0]))
    ax.append(plt.subplot(gs[2]))
    ax.append(plt.subplot(gs[3], sharey=ax[2], sharex=ax[2]))
    
    df['date_day'] = [datetime.datetime(d.year, d.month, d.day, 12, 0, 0,  tzinfo=dateutil.tz.tzlocal()) for d in df['date']]
    df['date_weekday'] = [d.weekday() for d in df['date_day']]
    df['date_hour'] = [d.hour+d.minute/60+d.second/3600 for d in df['date']]
    
    df_outgoing = df[df['sender_name'] == CFG_name_analyze]
    df_incoming = df[df['sender_name'] != CFG_name_analyze]
    
    # box plots
    for df_, ax_, title in zip((df_outgoing, df_incoming), (ax[0],ax[1]), (CFG_text_mailssent, CFG_text_mailsreceived)):
        df_d = df_.groupby('date_day').agg({'date_hour' : 'count', 'date_weekday': 'mean'})  # the mean of date_weekday does not change anything, they all should be the same

        avg = df_d['date_hour'].sum()/days_span  # compute average per day
        l = "%0.1f per day" % avg
        
        df_d = df_d.asfreq('D')  # we need to fill the missing dates to get the correct statistics
        df_d['date_hour'].fillna(0, inplace=True)  # counts will be 0
        df_d['date_weekday'] = [d.weekday() for d in df_d.index]  # we need to redo the weekday calculation for the filled values
        
        sns.barplot(x='date_weekday', y='date_hour', data=df_d, ci=95, order=np.arange(7), color="#2a7ab9", ax=ax_)  # palette="GnBu_d"
        
        ax_.axhline(y = avg, alpha=0.5, label=l, color="#303030", lw=1)
        ax_.text(5,avg+0.5, l, color="#303030", alpha=0.5)
        
        ax_.set_title(title)
        ax_.set_xticklabels([])
        ax_.set_ylabel("mails per day")
        ax_.set_xlabel("")
        ax_.yaxis.set_major_locator(mpl.ticker.MultipleLocator(2))

    # violin plots
    for df_, ax_, title in zip((df_outgoing, df_incoming), (ax[2],ax[3]), (CFG_text_mailssent, CFG_text_mailsreceived)):
        sns.violinplot(x="date_weekday", y="date_hour", data=df_, order=np.arange(7), color="#2a7ab9", scale="count", inner="sticks", bw=0.01, lw=1, ax=ax_)  #palette="GnBu_d"
        
        ax_.set_xlabel("")
        ax_.set_ylabel("hour")
        ax_.set_xticklabels(CFG_text_weekdays)
        ax_.set_ylim((0,24))
        ax_.yaxis.set_major_locator(mpl.ticker.MultipleLocator(4))
        ax_.yaxis.set_major_formatter(mpl.ticker.FormatStrFormatter('%0.0f:00'))
    fig.tight_layout() 
    
    
    
    
    
    
# main

if __name__ == '__main__':
    if CFG_offline_data and os.path.isfile( CFG_offline_data_filename ):
        headers = pickle.load( open( CFG_offline_data_filename, "rb" ) )
    else:
        print('Fetching emails from server...', end=' ')
        headers = getHeaders(CFG_mail_account, CFG_mail_passwd, CFG_imap_folder,CFG_imap_download_lastdays)
        if CFG_online_data_save: pickle.dump( headers, open( CFG_offline_data_filename, "wb" ) )
        print('done.')

    sns.set_context("notebook", font_scale=1.8, rc={"lines.linewidth": 2.5})
    
    print('Parsing...', end=' ')
    df = parseHeaders(headers)
    df = df[df['date'] > CFG_date_start]
    date_end = df['date'].max()
    date_start = df['date'].min()
    days_span = (date_end-date_start).days

    print('Plotting some statistics...')
    weekdayPlot(df.groupby('message_id_intern').head(1), days_span)

    if len(sys.argv)>1:
        if (sys.argv[1]=="-s") or (sys.argv[1]=="-save"):
            plt.savefig(os.path.splitext(os.path.basename(__file__))[0]+'.png', bbox_inches='tight', dpi=100)
        else:
            plt.show()
    else:
            plt.show()
    