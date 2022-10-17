#!/usr/bin/env python
#
# python/tk script to check email via IMAP
# runs under Linux (3.7) or Windows (3.10)
#
# <- Last updated: Fri Jul 29 19:37:05 2022 -> SGK
#
import email, getpass, imaplib, sys, os, re, platform
import tkinter as tk
from tkinter import ttk
from functools import partial
from email.header import decode_header
#
# ---------------------------------------------------------------------------
#
# global variables if defined here
wList     = {}
mboxProps = []
imgWidth  = 0
imgHeight = 0
#
if platform.system() == 'Windows':
    home = os.getenv('HOMEDRIVE')+os.getenv('HOMEPATH')
else:
    home = os.getenv('HOME')
#
# offsets ought to be fraction of image size
#
options = { 'rcFile'      : home+'/.ckmrc+',
#           'bitmapDir'   : 's1/',
#           'labelSep'    : ' ',
            'bitmapDir'   : 's2/',
            'labelSep'    : '\n',
#           'bitmapDir'   : 's3/',
#           'labelSep'    : '\n',
            'debug'       : 0,          ### 0, 1, 2
            'useEnv4Pass' : 0,
            'useCLI4Pass' : 0,
            'geometry'    : '-800+150', ### '+20+170' ## '-2+368'
            'passwdGeom'  : '+250+250',
            'checkTime'   : 3,
            'popUpWidth'  : 105,
            'popUpLeft'   : 1,          ### 0
            'soundFile'   : 'attention.wav',
            'useXBell'    : 0,
            'volume'      : 100,        ### %
            'useSounds'   : 1,          ### 0
            'noFork'      : 0,          ### 1
            'fakeIt'      : 0,          ### 1
}
#
# ------------------------------------------------------------------------
#
def CheckMailBox(mbox, debug = 0):
    #
    if options['fakeIt']:
        return(10, 3)
    #
    waitCursor = 'watch'
    wList['.'].config(cursor = waitCursor)
    wList['.'].update()
    #
    mailbox = mbox['mailbox']
    try:
        mailbox.select('Inbox', readonly = True)
    except:
        print('ckmail: CheckMailBox('+mbox['name']+') failed to select(Inbox)')
        return(-1, -1)
    #
    status, all = mailbox.search(None, "ALL") # charset, criterion[, ...]
    if status != 'OK':
        print('ckmail: CheckMailBox('+mbox['name']+') failed '+status)
        exit()
    all  = all[0].split()
    nMsg = len(all)
    status, new = mailbox.search(None, "UNSEEN") 
    new = new[0].split()
    ## print('>>new=', new)
    # save the list of new messages
    mbox['nList'] = new
    nNew  = len(new)
    if debug > 0:
        print('ckmail:', mbox['name'], 'has', nMsg, 'message(s),', nNew, 'new')
    #
    wList['.'].config(cursor = '')
    wList['.'].update()
    #
    return (nMsg, nNew)
#        
# ---------------------------------------------------------------------------
#
# does that widget exists (from wList and name)
def Exists(name):
    if name in wList:
        if wList[name].winfo_exists():
            ## print('>>',name,'exists')
            return True
    return False
#
# ------------------------------------------------------------------------
#
def WhenOK(popW):
    popW.destroy()
#
# ------------------------------------------------------------------------
#
def WhenSelect(wName, e):
    #
    w = e.widget
    sel = w.curselection()
    nSel = len(sel)
    # place holders so far
    if nSel == 0:
        state = 'disabled'
        ## print('>> WhenSelect(): nothing selected')
    else:
        state = 'active'
        ## print('>> WhenSelect():',nSel,'item(s) selected')
        ##for index in sel:
        ##    value = w.get(index)
        ##    print('>>   index %d: "%s"' % (index, value))
    #
    w = wList[wName+'.read']
    w.configure(state = state)
    w = wList[wName+'.del']
    w.configure(state = state)
#
# ------------------------------------------------------------------------
#
def WhenToggleSelection(i, wName, debug):
    #
    if debug > 0:
        print('ckmail: WhenToggleSelection('+str(i)+', '+wName+')')
    w = wList[wName+'.listbox']
    sel = w.curselection()
    all = w.get(0, 'end')
    nSel = len(sel)
    nAll = len(all)
    #
    for i in range(nAll):
        if i in sel:
            w.select_clear(i)
        else:
            w.select_set(i)
    #
    if nAll == nSel:
        state = 'disabled'
    else:
        state = 'active'
    w = wList[wName+'.read']
    w.configure(state = state)
    w = wList[wName+'.del']
    w.configure(state = state)
#
# ------------------------------------------------------------------------
#
def WhenDelMR(i, wName, mbox, delete, debug):
    # 
    if debug > 1:
        print('ckmail: WhenDelMR('+str(i)+', '+wName+', '+str(delete)+')')
    mailbox = mbox['mailbox']
    try:
        mailbox.select('Inbox', readonly = False)
        ok = 1
    except:
        print('ckmail: WhenDelMR('+mbox['name']+') failed to select(Inbox)')
        ok = 0
    if ok == 1:
        listBoxW = wList[wName+'.listbox']
        sel = listBoxW.curselection()
        nList = mbox['nList']
        for index in sel:
            num = nList[index]
            ## print('>>num=', num)
            if delete:
                flag =  '\\Deleted'
                txt = 'deleting'
            else:
                flag =  '\\Seen'
                txt = 'marking as read'
            #
            value = listBoxW.get(index)
            if debug > 0:
                print('ckmail: WhenDelMR('+mbox['name']+') %s msg# %d: "%s"' % (txt, index, value))
            mailbox.store(num, '+FLAGS', flag)
        #
        mailbox.expunge()
    # 
    popW = wList[wName]
    popW.destroy()
    #
    WhenCheckForMail(i, mbox)
#
# ------------------------------------------------------------------------
#
def PopMsgList(i, mbox, mainW, text):
    #
    name  = mbox['name']
    color = mbox['color']
    mDel  = int(mbox['mDel'])
    nText = len(text)
    wName = 'ckpop.'+name
    debug = options['debug']
    #
    if debug > 0:
        print('ckmail: PopMsgList('+name+')', nText, 'line(s)')
        #
    if Exists(wName):
        if debug > 1:
            print('ckmail: PopMsgList()', wName, 'exists')
        #
        textW = wList[wName]
        if nText == 0:
            textW.destroy()
            return
        else:
            #
            hx = 10
            if nText < hx:
                hx = nText
            #
            listboxW = wList[wName+'.listbox']
            h = listboxW.cget('height')
            # if big enough
            if hx <= h:
                listboxW.configure(listvariable = tk.StringVar(value=text))
##              listboxW.delete('0', 'end')
##              listboxW.insert('0', text)
                listboxW.see('end')
                return
            else:
                textW.destroy() 
                if debug > 1:
                    print('ckmail: PopMsgList()', wName, 'too small')
    #
    if nText > 0:
        #
        ## no rootx() rooty()
        x = mainW.winfo_x()
        y = mainW.winfo_y()
        ## print('>>x, y, imgWidth, imgHeight =', x, y, imgWidth, imgHeight)
        #
        # font size -> 8/9
        fsz = 8
        if platform.system() != 'Windows':
            fsz = 9
        #
        if options['popUpLeft'] == 1:
            x -= fsz*options['popUpWidth'] - imgWidth*i - 20
            y += 50 + int(imgHeight/2*i)
        else:
            x += 50 + imgWidth*i
            y += 50 + int(imgHeight/2*i)
        #
        h = 10
        if nText < h:
            h = nText
        #
        geom ='+'+str(x)+'+'+str(y)
        ## print('>>geom=', geom)
        #
        popW = tk.Toplevel()
        string = 'CkM: new mail ('+name+')'
        popW.title(string)
        popW.iconname(string)
        popW.geometry(geom)
        popW.transient(mainW)
        popW.resizable(0, 0)
        #
        wList[wName] = popW
        #
        s = 's'
        if nText == 1:
            s = ''
        #
        labelText = str(nText)+' new message'+s+' in '+name
        labelW = tk.Label(popW,
                          text = labelText,
                          background = '#f0f0f0',
                          foreground = color)
        frameW = tk.Frame(popW)
        font = ('Courier', 
                10, 
                'bold')
        selectmode = 'multiple'  ### 'extended' # diff??
        listboxW = tk.Listbox(frameW,
                              listvariable = tk.StringVar(value=text),
                              selectmode = selectmode,
                              setgrid    = 1,
                              height     = h,
                              font       = font,
                              width      = options['popUpWidth'])

        wList[wName+'.listbox'] = listboxW
        scrollbarW = ttk.Scrollbar(frameW,
                                   orient = 'vertical',
                                   command = listboxW.yview)
        listboxW['yscrollcommand'] = scrollbarW.set
        listboxW.see('end')
        #
        if mDel == 1:
            listboxW.bind('<<ListboxSelect>>', partial(WhenSelect, wName))
        #
        buttonsHolderW = tk.Frame(popW) 
        okW = tk.Button(buttonsHolderW,
                        text        = 'OK',
                        background  = '#f0f0f0',
                        foreground  = color,
                        command     = partial(WhenOK, popW))
        delW = tk.Button(buttonsHolderW,
                         text       = 'Delete',
                         state      = 'disabled',
                         background = '#f0f0f0',
                         foreground = color,
                         command    = partial(WhenDelMR, i, wName, mbox, 1, debug) )
        readW = tk.Button(buttonsHolderW,
                          text       = 'Mark as read',
                          state      = 'disabled',
                          background = '#f0f0f0',
                          foreground = color,
                          command    = partial(WhenDelMR, i, wName, mbox, 0, debug))
        if mDel == 1:
            state = 'active'
        else:
            state = 'disabled'

        togW = tk.Button(buttonsHolderW,
                         text       = 'Toggle',
                         state      = state,
                         background = '#f0f0f0',
                         foreground = color,
                         command = partial(WhenToggleSelection, i, wName, debug))
        #
        labelW.  pack(side = 'top',  fill = 'x') 
        frameW.  pack(side = 'top',  fill = 'x') 
        listboxW.pack(side = 'left', fill = 'x') 
        scrollbarW.pack(side = 'right',   fill = 'y')
        buttonsHolderW.pack(side = 'top', fill = 'x') 
        okW.  pack(side = 'left',  fill = 'x', expand = 1) 
        delW. pack(side = 'right', fill = 'x', expand = 0) 
        readW.pack(side = 'right', fill = 'x', expand = 0) 
        togW. pack(side = 'right', fill = 'x', expand = 0)
        #
        wList[wName+'.read'] = readW
        wList[wName+'.del']  = delW
#
# ------------------------------------------------------------------------
#
def WhenToggleOpt(o, w):
    #
    v = w.get()
    v = 1 - int(v)
    if v < 0:
        v = 0
    ## print('>> w,o,v=', w, o, v)
    w.delete(0, 'end')
    w.insert(0, str(v))
#
# ---------------------------------------------------------------------------
#
def WhenCancelOpts():
    #
    #
    nMBox = len(mailBoxes)
    for i in range(nMBox):
        w = wList['.icons.mailbox'+str(i)]
        w.configure(state = 'active')
    #
    wList['.icons.optsButton'].configure( state = 'active')
    wList['.icons.checkButton'].configure(state = 'active')
    ## wList['.'].Unbusy()
    wList['.'].config(cursor='')
    wList['.'].update()
    #
    wList['.ckoptns'].destroy()
#
# ---------------------------------------------------------------------------
#
def WhenDoneOpts(wE):
    #
    debug = options['debug']
    for k in wE:
        v = wE[k].get()
        if debug > 1:
            print('ckmail: WhenDoneOpts() options['+k+'] =', v)
        if isinstance(options[k], int):
            options[k] = int(v)
        else:
            options[k] = v
    #
    WhenCancelOpts()
#
# ---------------------------------------------------------------------------
#
def PopEditOptions():
    #
    ## wList['.'].Busy()
    waitCursor = 'watch' # wait
    wList['.'].config(cursor = waitCursor)
    wList['.'].update()
    #
    nMBox = len(mailBoxes)
    for i in range(nMBox):
        w = wList['.icons.mailbox'+str(i)]
        w.configure(state = 'disabled')
    #
    wList['.icons.optsButton'].configure( state = 'disabled')
    wList['.icons.checkButton'].configure(state = 'disabled')
    #
    w = wList['.']
    # not rootx rooty 
    x = w.winfo_x()
    y = w.winfo_y()
    if options['popUpLeft'] == 1:
        x -= imgWidth*2+55
    else:
        x += imgWidth*2
    y += int(imgHeight/2)
    #
    geom = '+'+str(x)+'+'+str(y)
    string = 'CkM: options'
    popW = tk.Toplevel()
    popW.title(string)
    popW.iconname(string)
    popW.geometry(geom)
    popW.transient(wList['.'])
    popW.resizable(0, 0)
    # catch the close on clicking 'X', equiv to cancel
    popW.protocol('WM_DELETE_WINDOW', WhenCancelOpts)
    #
    wList['.ckoptns'] = popW
    #
    opts = {'1:Check interval:min'   : 'checkTime',
            '2:Pop width:pxl'        : 'popUpWidth',
            '3:Pop on left side:0|1' : 'popUpLeft',
            '4:Sound on:0|1'         : 'useSounds',
            '5:Volume:%'             : 'volume',
            '6:Sound file:-'         : 'soundFile',
            '7:Debug level:0|1'      : 'debug',
    }
    #
    wEList = {}
    #
    frameW1 = tk.Frame(popW)
    for key in opts:
        frameW2 = tk.Frame(frameW1,
                           border = 1,
                           relief = 'ridge' )
        #
        words = key.split(':')
        labelW = tk.Label(frameW2,
                          text = words[1]+' : ')
        labelW.pack(side = 'left')
        #
        entryW = tk.Entry(frameW2)
        entryW.pack(side = 'left',
                    fill = 'x', 
                    expand = 1)
        entryW.insert('end', options[opts[key]])
        #
        if words[2] == '0|1':
            ## wE -> configure(-state = 'disabled')
            xW = tk.Button(frameW2,
                           text = '0|1',
                           command = partial(WhenToggleOpt, opts[key], entryW))
        else:
            xW = tk.Label(frameW2,
                          text = words[2])
        xW.pack(side = 'right')
        #
        wEList[opts[key]] = entryW
        #
        frameW2.pack(side = 'top', fill = 'x', expand = 1)
        
    frameW1.pack(side = 'top', fill = 'x', expand = 1)
    frameW2 = tk.Frame(frameW1, border = 2)
    buttonW1 = tk.Button(frameW2,
                         text = 'Cancel',
                         command = WhenCancelOpts)
    buttonW2 = tk.Button(frameW2,
                         text = 'Done',
                         command = partial(WhenDoneOpts, wEList))
    frameW2.pack(side = 'bottom', fill = 'x', expand = 1)
    buttonW1.pack(side = 'left',   fill = 'x', expand = 1)
    buttonW2.pack(side = 'right',  fill = 'x', expand = 1)
#        
# ---------------------------------------------------------------------------
#
def GetNewMessages(mbox, debug = 0):
    #
    if debug > 0:
        print('ckmail: GetNewMessages('+mbox['name']+')')
    if options['fakeIt']:
        return ['fake line 1',
                'fake line 2',
                'fake line 3']
    #
    waitCursor = 'watch' # wait
    wList['.'].config(cursor = waitCursor)
    wList['.'].update()
    #
    mailbox = mbox['mailbox']
    status, new = mailbox.search(None, "UNSEEN") 
    new = new[0].split() 
    nNew  = len(new)
    list = []
    for item in new:
        #
        status, data = mailbox.fetch(item, "(RFC822.HEADER)")
        if status == 'OK':
            email_header = email.message_from_bytes(data[0][1])
            header = {}
            for keyword in ['From', 'Subject', 'Date']:
                hdrVal = ''
                hdrList = decode_header(email_header[keyword])
                for entry in hdrList:
                    string, encoding = entry
                    if isinstance(string, bytes):
                        if encoding == None:
                            encoding = 'utf-8'
                        string = string.decode(encoding)
                    hdrVal = hdrVal + string
                header[keyword] = hdrVal
                ## print('>> '+keyword+':', hdrVal)
            #
            if debug > 1:
                print ("Message %s %s %s %s" % (str(item, 'UTF-8'),
                                                header['From'],
                                                header['Date'],
                                                header['Subject']))
            #
            header['From'] = re.sub('<.*>', '', header['From'])
            header['From'] = re.sub('"',    '', header['From'])
            header['From'] = re.sub("'",    '', header['From'])
            header['From'] = re.sub('^ *',  '', header['From'])
            header['From'] = re.sub(' *$',  '', header['From'])
            ## microsoft add \n\r
            header['From'] = re.sub('\n',   '', header['From'])
            header['From'] = re.sub('\r',   '', header['From'])
            ## print(' >>'+header['From']+'<')
            #
            ## print(' >>'+header['Date']+'<')
            header['Date'] = re.sub('^..., ', '', header['Date'])
            header['Date'] = header['Date'].strip()
            header['Date'] = re.sub('  ',    ' ', header['Date'])
            w = header['Date'].split(' ')
            ## print(','.join(w))
            header['Date'] = w[0]+' '+w[1]+' '+w[3]
            ## print('>>>'+header['Date']+'<')
            #
            header['Subject'] = re.sub('\n *', ' ', header['Subject'])
            header['Subject'] = re.sub('\f *', ' ', header['Subject'])
            header['Subject'] = re.sub('\r *', ' ', header['Subject'])
            # convert subject to ascii only (b/c emoji)
            header['Subject'] = header['Subject'].encode('ascii', 'replace').decode()
            #
            str_header = "%-30.30s %16.16s %.64s" % (header['From'],
                                                     header['Date'],
                                                     header['Subject'])
            #
            list.append(str_header)
    #
    wList['.'].config(cursor = '')
    wList['.'].update()
    #
    return list
#        
# ---------------------------------------------------------------------------
#
def PlaySound(soundFile, volume):
    #
    ## print('>> PlaySound(soundFile, volume):', soundFile, volume)
    if platform.system() == 'Linux':
        pc = volume/100
        cmd = 'play -q -v '+str(pc)+' '+soundFile
        if options['debug'] > 1:
            print('ckmail: PlaySound(): executing '+cmd)
        os.system(cmd)
    else:
        # pip install playsound==1.2.2
        # no volume control, tho
        from playsound import playsound
        playsound(soundFile, False)
#
# ---------------------------------------------------------------------------
#
def CheckAllMail(infos, oob = False):
    #
    # oob: out of band check: do not set alarm
    #
    mailBoxes = infos['mailBoxes']
    mboxProps = infos['mboxProps']
    images    = infos['images']
    #
    soundFile = options['soundFile']
    volume    = options['volume']
    delay     = options['checkTime']*60000
    useXBell  = options['useXBell']
    useSounds = options['useSounds']
    debug     = options['debug']
    #
    mainW     = wList['.']
    #
    if not oob:
        mainW.after(delay, lambda: CheckAllMail(infos))
    #
    nMBox = len(mailBoxes)
    if options['debug'] != 0:
        print('ckmail: CheckAllMail() '+str(nMBox)+' mailbox(es)')
    #
    nNewTot = 0
    for i in range(nMBox):
        #
        mbox = mailBoxes[i]
        #
        # previous nList
        pNList = mbox['nList']
        #
        nMsg, nNew = CheckMailBox(mbox, debug = debug)
        if nMsg == -1 and nNew == -1:
            if options['debug'] != 0:
                print('calling GetMailBox(mboxProp)')
            mboxProp = mboxProps[i]
            mbox = GetMailBox(mboxProp, debug = debug)
            nMsg, nNew = CheckMailBox(mbox, debug = debug)
            if nMsg == -1 and nNew == -1:
                exit()
            else:
                mailBoxes[i] = mbox  
        #
        mbox['nMsg'] = nMsg
        mbox['nNew'] = nNew
        #
        labelW = wList['.icons.label'  +str(i)]
        iconW  = wList['.icons.mailbox'+str(i)]
        #
        text = mbox['name']+options['labelSep']+'('+str(nMsg)+'/'+str(nNew)+')'
        labelW.configure(text = text)
        #
        if nNew == 0:
            image = images['empty']
        else:
            image = images['full']
        iconW.configure(image = image)
        #
        if nNew > 0:
            #
            list = GetNewMessages(mbox, debug = debug)
            PopMsgList(i, mbox, mainW, list)
            #
            # current nList
            nList = mbox['nList']
            ##print('>>pNList:', pNList)
            ##print('>>nList: ', nList)
            #
            # check if new
            for entry in nList:
                if entry in pNList:
                    junk = 0 # noop
                    ## print('>>', entry,' is not new')
                else:
                    nNewTot = nNewTot + 1
        else:
            # destroy if exists
            name  = mbox['name']
            wName = 'ckpop.'+name
            if Exists(wName):
                popW = wList[wName]
                popW.destroy()
    #
    ## print('>>>', nNewTot)
    if nNewTot > 0:
        if useSounds != 0:
            if useXBell == 1:
                print("\g")
            else:
                PlaySound(soundFile, volume)
#
# ---------------------------------------------------------------------------
#
def WhenCheckAllMail(infos):
    # not (event):
    if options['debug'] > 0:
        print('ckmail: WhenCheckAllMail()')
    CheckAllMail(infos, oob = 'True')
#
# ---------------------------------------------------------------------------
#
def WhenCheckForMail(i, mbox):
    # not (event):
    if options['debug'] > 0:
        print('ckmail: WhenCheckForMail('+mbox['name']+')')
    nMsg, nNew = CheckMailBox(mbox, debug = options['debug'])
    if nNew > 0:
        #
        list = GetNewMessages(mbox, debug = options['debug'])
        PopMsgList(i, mbox, mainW, list) 
    else:
        #
        name  = mbox['name']
        wName = 'ckpop.'+name
        if Exists(wName):
            popW = wList[wName]
            popW.destroy()
    #
    labelW = wList['.icons.label'  +str(i)]
    iconW  = wList['.icons.mailbox'+str(i)]
    #
    text = mbox['name']+options['labelSep']+'('+str(nMsg)+'/'+str(nNew)+')'
    labelW.configure(text = text)
    #
    if nNew == 0:
        image = images['empty']
    else:
        image = images['full']
    iconW.configure(image = image)        
#
# ---------------------------------------------------------------------------
#
def WhenEditOptions():
    if options['debug'] > 1:
        print('ckmail: WhenEditOptions()')
    PopEditOptions()
#
# ------------------------------------------------------------------------
#
def WhenToggleSound():
    if options['debug'] > 1:
        print('ckmail: WhenToggleSound() -> ', options['useSounds'])
    #
    options['useSounds'] = 1 - options['useSounds'];
    #
    if options['useSounds'] == 0:
        image = images['soundOff']
    else:
        image = images['soundOn']
    #
    iconW = wList['.icons.soundButton']
    iconW.configure(image = image)
    if options['debug'] > 1:
        print('ckmail: WhenToggleSound() -> ', options['useSounds'])
#
# ---------------------------------------------------------------------------
#
def WhenDone():
    exit()
#
# ---------------------------------------------------------------------------
#
def GetMailBox(prop, debug = 0):
    #
    if debug > 0:
        print('ckmail: '+prop['name']+': logging as', prop['email'])
    #
    if options['fakeIt']:
        mailbox = None
    else:
        #
        mailbox = imaplib.IMAP4_SSL(prop['host'], prop['port'])
        try:
            mailbox.login(prop['email'], prop['passwd'])
        except:
            if options['useCLI4Pass'] == 1:
                print('ckmail: failed to login as', prop['email'])
                exit()
            if options['debug'] > 1:
                print('ckmail: failed to login as', prop['email'])
            return 'error'
        #
    if debug > 0:
        print()
    #
    nMsg = 0
    nNew = 0
    mbox = { 'mailbox': mailbox,
             'name'   : prop['name'],
             'color'  : prop['color'],
             'mDel'   : prop['mDel'],
             'nMsg'   : nMsg, 'nNew': nNew}
    return mbox
#
# ---------------------------------------------------------------------------
#
def BuildGUI(mailBoxes, mboxProps):
    #
    global imgWidth
    global imgHeight
    #
    nMBox = len(mailBoxes)
    #
    mainW = tk.Tk()
    ##  mainW.optionAdd('*font', options['font'])
    #
    bDir = options['bitmapDir']
    bitmaps = {'empty': bDir+'empty.png' ,
               'full' : bDir+'full.png' ,
               'ico'  : bDir+'mbox.png',
               'quit' : bDir+'quit.png',
               'check': bDir+'check.png',
               'opts' : bDir+'opts.png',
               'soundOn'  : bDir+'soundOn.png',
               'soundOff' : bDir+'soundOff.png',
    }
    #
    images = {}
    for key in bitmaps:
        images[key] = tk.PhotoImage(file = bitmaps[key])
        if imgWidth == 0:
            imgWidth  = images[key].width()
            imgHeight = images[key].height()
    #
    iconsW = tk.Frame(mainW, relief = 'raised', borderwidth = 1)
    #
    for i in range(nMBox):
        mbox = mailBoxes[i]
        if (mbox['nNew'] == 0):
            image = images['empty']
        else:
            image = images['full']
            #
        holderW = tk.Frame(master = iconsW, background = '#f0f0f0')
        iconMailboxW = tk.Button(master = holderW,
                                 image  = image,
                                 background = mbox['color'],
                                 activebackground = 'pink',
                                 relief = 'raised',
                                 command = partial(WhenCheckForMail, i, mbox))
        #
        text = mbox['name']+options['labelSep']+'('+str(mbox['nMsg'])+'/'+str(mbox['nNew'])+')'
        labelW = tk.Label(master = holderW,
                          text = text,
                          background = '#f0f0f0',
                          foreground = mbox['color'])
        iconMailboxW.pack()
        labelW.pack()
        holderW.pack(side = 'left')
        #
        wList['.icons.label'  +str(i)] = labelW
        wList['.icons.mailbox'+str(i)] = iconMailboxW
        #
    #
    iconsW.pack(side = 'left')
    iconsW = tk.Frame(mainW, relief = 'raised', borderwidth = 1)
    #
    iconsQuitButtonW  = tk.Button(iconsW,
                                  image = images['quit'],
                                  #text = 'Quit',
                                  command = WhenDone)  
    #
    iconsCheckButtonW = tk.Button(iconsW,
                                  image = images['check'],
                                  #text = 'Check',
                                  )
    #
    if options['useSounds'] == 1:
        image = images['soundOn']
    else:
        image = images['soundOff']
    iconsSoundButtonW  = tk.Button(iconsW,
                                   image = image,
                                   #text = 'SoundOnOff',
                                   command = WhenToggleSound)
    #
    iconsOptsButtonW  = tk.Button(iconsW,
                                  image = images['opts'],
                                  #text = 'Opts',
                                  command = WhenEditOptions)
    #
#   iconsXxxW         = tk.Label(iconsW,
#                                 # image = images['box'],
#                                 text = ' ')
    iconsQuitButtonW.pack (side = 'top', fill = 'x')
    iconsCheckButtonW.pack(side = 'top', fill = 'x')
    iconsSoundButtonW.pack(side = 'top', fill = 'x')
    iconsOptsButtonW.pack (side = 'top', fill = 'x')
#   iconsXxxW.pack(side = 'bottom', fill = 'both')
    iconsW.pack(side = 'left', fill = 'y')
    #
    wList['.icons.quitButton']  = iconsQuitButtonW
    wList['.icons.optsButton']  = iconsOptsButtonW
    wList['.icons.soundButton'] = iconsSoundButtonW
    wList['.icons.checkButton'] = iconsCheckButtonW
    #
    mainW.title('CkM')
    mainW.iconname('CkM')
    mainW.iconphoto(True, images['ico'])
    mainW.resizable(0, 0)
    mainW.geometry(options['geometry'])
    #
    infos = {'mailBoxes'  : mailBoxes,
             'mboxProps'  : mboxProps,
             'images'     : images,
             }
    #
    iconsCheckButtonW.configure(command = partial(WhenCheckAllMail, infos))
    #
    # check in 1 sec
    mainW.after(1000, lambda: CheckAllMail(infos))
    #
    # the mainloop() can't be outside this routine,
    # unless return ptr to images so they don't get destroyed
    return(mainW, images)
#
# --------------------------------------------------------------------------
#
def ParseArgs(args, optsIn, rcOnly = 0):
    #
    USAGE = ['', 'usage: ckmail [options]',
             '', 'where options:',
             '  -d -dd -ddd:          debug level',
             '  -noFork               do not fork',
             '  -useE4P               use env for passwd',
             '  -useC4P               use cli for passwd',
             '  -noSound              no sound',
             '  -useXBell             no sound, but X-bell',
             '  -rcFile <file>        config fileName', 
             '  -geom +X+Y            location',
             '  -popUpWidth           popUp width',
             '  -popUpRight           popUp to the right',
             '  -checkTime <time>     check time interval, in min',
             '  -soundFile <fileName> ', 
             '  -volume NNN           volume (NNN = 1-100 or in %)',
             '',
             '  -help              show this help',
             ' Ver 1.2/1']
    #
    vars = ['-d -dd -ddd -noFork -useE4P -useC4P -noSound -useXBell -rcFile ',
            '-geom -popUpWidth -popUpRight -checkTime -soundFile -volume -help']
    vars = ' '.join(vars).split(' ')
    #
    options = {}
    for key in optsIn:
        options[key] = optsIn[key]
    #
    nErr = 0
    showHelp = 0
    i = 1
    while i < len(args):
        arg = args[i]
        if arg in vars:
            if arg == '-d':
                options['debug'] = 1
            if arg == '-dd':
                options['debug'] = 2
            if arg == '-ddd':
                options['debug'] = 3
            if arg == '-noFork':
                options['noFork'] = 1
            if arg == '-useE4P':
                options['useEnv4Pass'] = 1
            if arg == '-useC4P':
                options['useCLI4Pass'] = 1
            if arg == '-noSound':
                options['useSounds'] = 0
            if arg == '-useXBell':
                options['useXBell'] = 1
            if arg == '-rcFile':
                i += 1
                if i == len(args):
                    print('ckmail: missing argument to', arg)
                    nErr += 1
                else:
                    options['rcFile'] = args[i]
            if arg == '-geom':
                i += 1
                if i == len(args):
                    print('ckmail: missing argument to', arg)
                    nErr += 1
                else:
                    options['geometry'] = args[i]
            if arg == '-popUpWidth':
                i += 1
                if i == len(args):
                    print('ckmail: missing argument to', arg)
                    nErr += 1
                else:
                    options['popUpWidth'] = int(args[i])
            if arg == '-popUpRight':
                options['popUpLeft'] = 0
            if arg == '-checkTime':
                i += 1
                if i == len(args):
                    print('ckmail: missing argument to', arg)
                    nErr += 1
                else:
                    options['checkTime'] =  int(args[i])
            if arg == '-soundFile':
                i += 1
                if i == len(args):
                    print('ckmail: missing argument to', arg)
                    nErr += 1
                else:
                    options['soundFile'] = args[i]
            if arg == '-volume':
                i += 1
                if i == len(args):
                    print('ckmail: missing argument to', arg)
                    nErr += 1
                else:
                    options['volume'] = int(args[i])
            if arg == '-help':
                showHelp = 1
        else:
            print('ckmail: invalid arg "'+arg+'"')
            nErr += 1
        i += 1
        #
    #
    if nErr > 0 or showHelp == 1:
        print("\n".join(USAGE))
        exit()
    #
    if rcOnly:
        rcFile = options['rcFile']
        debug  = options['debug']
        options = {}
        for key in optsIn:
            options[key] = optsIn[key]
        options['rcFile'] = rcFile
        options['debug']  = debug
    #
    return options
#
# ---------------------------------------------------------------------------
#
def ParseRCFile(optsIn):
    #
    options = {}
    for key in optsIn:
        options[key] = optsIn[key]
    #
    try:
        file = open(options['rcFile'], 'r')
    except:
##        if options['debug'] > 0:
        print('ckmail: ParseRCFile() rcFile "'+options['rcFile']+'" not found.')
        return options
    #
    i = 0
    setProp = 0
    for line in file:
        i += 1
        ## print('  >>'+line.strip()+'<<')
        line = line.strip()
        if not re.search('^!', line):
            if line == '{':
                if setProp == 0:
                    setProp = 1
                    prop = {}
                else:
                    print('ckmail: ParseRCFile() ignoring line #', i, ':', line)
            elif line == '}':
                if setProp == 1:
                    setProp = 0
                    mboxProps.append(prop)
                else:
                    print('ckmail: ParseRCFile() ignoring line #', i, ':', line)
            elif re.search('=', line):
                w = line.split('=')
                key = w[0].strip()
                val = w[1].strip()
                if setProp == 1:
                    prop[key] = val
                else:
                    if key in options:
                        if isinstance(options[key], int):
                            options[key] = int(val)
                        else:
                            options[key] = val
                    else:
                        print('ckmail: ParseRCFile() ignoring line #', i, ':', line)
                        ##print('>> v>>'+val+'<<')
                        ##print('>> k>>'+key+'<<')

            elif line != '':
                print('ckmail: ParseRCFile() ignoring line #', i, ':', '>'+line+'<')
            #
    file.close()
    #
    return options
#
# ---------------------------------------------------------------------------
#
def WhenShowPasswd(buttonW, entryW):
    # show/hide passwd entried
    text = buttonW.cget('text')
    if text == 'show':
        entryW.configure(show = '')
        buttonW.configure(text = 'hide')
    else:
        entryW.configure(show  = '*')
        buttonW.configure(text = 'show')
#
# ---------------------------------------------------------------------------
#
def WhenCancelPasswd():
    exit()
#
# ---------------------------------------------------------------------------
#
def WhenOKPasswd():
    #
    i = 0 
    for mboxProp in mboxProps:
        entryW = wList['.passwd.entry'+str(i)]
        passwd = entryW.get()
        mboxProp['passwd'] = passwd
        i += 1
    #
    popW =  wList['.passwd']
    popW.destroy()
#
# ------------------------------------------------------------------------
#
def PopWrongPasswd(cli = 0):
    popW = tk.Tk()
    string = 'CkM: Wrong Password(s)'
    popW.title(string)
    popW.iconname(string)
    geom = options['passwdGeom']
    popW.geometry(geom)
    popW.resizable(0, 0)
    popW.protocol('WM_DELETE_WINDOW', WhenCancelPasswd)
    wList['.passwd'] = popW
    #
    frameW = tk.Frame(popW)
    labelW = tk.Label(frameW, text = 'ckmail: login(s) failed!', width = 20)
    quitW = tk.Button(frameW, text = 'Quit', command = WhenCancelPasswd)
    labelW.pack(side  = 'top', fill = 'x', expand = 1)
    quitW.pack(side = 'bottom',  fill = 'x', expand = 1)
    frameW.pack(side = 'top', fill = 'x', expand = 1)
    #
    popW.mainloop()
    #
    return
#
# ------------------------------------------------------------------------
#
def PopGetPasswd(cli = 0):
    #
    if cli == 1:
        for mboxProp in mboxProps:
            try:
                passwd = getpass.getpass('ckmail: enter password for '+mboxProp['name']+': ')
            except:
                print()
                exit()
            mboxProp['passwd'] = passwd
        return
    #
    popW = tk.Tk()
    string = 'CkM: Enter Password(s)'
    popW.title(string)
    popW.iconname(string)
    geom = options['passwdGeom']
    popW.geometry(geom)
    popW.resizable(0, 0)
    popW.protocol('WM_DELETE_WINDOW', WhenCancelPasswd)
    wList['.passwd'] = popW
    i = 0 
    for mboxProp in mboxProps:
        frameW = tk.Frame(popW)
        labelW = tk.Label(frameW, text = 'Password for '+mboxProp['name']+': ', width = 20)
        entryW = tk.Entry(frameW, show = '*')
        try:
            text = mboxProp['passwd']
        except:
            text = ''
        entryW.insert(0, text)
        buttonW = tk.Button(frameW, text = 'show', width = 5)
        buttonW.configure(command=partial(WhenShowPasswd, buttonW, entryW))
        labelW.pack(side  = 'left')
        entryW.pack(side  = 'left', fill = 'x', expand = 1)
        buttonW.pack(side = 'right')
        frameW.pack(side  = 'top')
        wList['.passwd.entry'+str(i)] = entryW
        i += 1
    #
    frameW = tk.Frame(popW)
    okW = tk.Button(frameW, text = 'OK', command = WhenOKPasswd)
    quitW = tk.Button(frameW, text = 'Quit', width = 5, command = WhenCancelPasswd)
    okW.pack(  side = 'left', fill = 'x', expand = 1)
    quitW.pack(side = 'right',  fill = 'x', expand = 0)
    frameW.pack(side = 'bottom', fill = 'x', expand = 1)
    #
    popW.mainloop()
    #
    return
#
# --------------------------------------------------------------------------
#
# os.name != 'nt'
# sys.platform != 'win32'
if platform.system() != 'Windows':
    if re.search('^localhost:', os.getenv('DISPLAY', 'localhost:')):
        options['useSounds'] = 0
#
options = ParseArgs(sys.argv, options, rcOnly = 1)
options = ParseRCFile(options)
options = ParseArgs(sys.argv, options)
#
if len(mboxProps) == 0:
    print('ckmail: #(mbox) == 0 - exiting.')
    exit()
    
if options['debug'] > 2:
    for key in options:
        print('ckmail: options['+key+'] =',options[key])
    for mboxProp in mboxProps:
        for key in mboxProp:
            print('ckmail: mbox['+key+'] =',mboxProp[key])
#
if options['fakeIt']:
    print('ckmail: fakeIt ==', options['fakeIt'])
#
mailBoxes = []
if options['useEnv4Pass']:
    for mboxProp in mboxProps:
        passwd = os.getenv('PASSWD_'+mboxProp['name'])
        ## print('>>PASSWD_'+mboxProp['name']+' >>'+str(passwd))
        mboxProp['passwd'] = passwd
else:
    PopGetPasswd(cli = options['useCLI4Pass'])
#
nErr = 0
for mboxProp in mboxProps:
    mbox = GetMailBox(mboxProp, debug = options['debug'])
    if mbox == 'error':
        nErr += 1
    else:
        mailBoxes.append(mbox)
    mbox['nList'] = []

if nErr > 0:
    if options['useCLI4Pass'] == 1:
        exit()
    else:
        PopWrongPasswd()
#
# fork is not available under Windows
if platform.system() != 'Windows':
    if options['noFork'] == 0:
        pid = os.fork()
        if pid > 0:
            exit()
#
mainW, images = BuildGUI(mailBoxes, mboxProps)
wList['.'] = mainW
mainW.mainloop()
#
