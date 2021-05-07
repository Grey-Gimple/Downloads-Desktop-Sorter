# required modules
import os
import shutil
import time
import subprocess
from datetime import date
from pathlib import Path

# file type to folder map
file_type_folder = {
    'audio':'Media\\Audio',
    'video':'Media\\Video',
    'image':'Media\\Image',
    'comp_archive':'File\\Comp&Archive',
    'font':'File\\Font',
    'executables':'File\\Executables',
    'system':'File\\System',
    'document':'File\\Document',
    'webpage':'Program\\Webpage',
    'db':'Program\\Database',
    'program':'Program\\Program'
}

# file extention to file type map
file_ext_file_type = {
    # Media
    'audio':('.8svx','.16svx','.aiff','.au','.bwf','.cdda','.dsf','.dff','.wav','.ra','.rm','.flac','.la','.pac','.ape','.ofr','.ofs','.off','.rka','.rkau','.shn','.tak','.thd','.tta','.wv','.wma','.brstm','.dts','.dtshd','.dtsma','.ast','.aw','.psf','.ac3','.amr','.mp1','.mp2','.mp3','.mpeg','.spx','.gsm','.aac','.mpc','.vqf','.ots','.swa','.vox','.voc','.dwd','.smp','.ogg','.mod','.mt2','.s3m','.xm','.it','.nsf','.mid','.ftm','.abc','.darms','.etf','.gp','.kern','.ly','.mei','.mus','.musx','.mxl','.mscx','.mxcz','.smdl','.sib','.niff','.ptb','.cust','.gym','.jam','.rmj','.sid','.spc','.txm',',vgm','.ym','.pvd'),
    'video':('.webm','.mkv','.flv','.vob','.ogv','.drc','.mng','.mts','.m2ts','.ts','.mov','.qt','.wmv','.yuv','.rmvb','.asf','.amv','.mp4','.m4p','.mpg','.mpv','.svi','.3pg','.mxf','.roq','.nsv','.f4v','.f4a','.f4b'),
    'image':('.jfif','.act','.ase','.gpl','.pal','.icc','.icm','.art','.blp','.bmp','.bti','.cd5','.cit','.cr2','.clip','.dds','.dib','.djvu','.exif','.gif','.gifv','.grf','.iff','.jng','.jpg','.jpeg', '.jp2','.jps','.kra','.lbm','.max','.miff','.msp','.nitf','.otb','.pbm','.pc1','.pc2','.pc3','.pcx','.pdn','.pgm','.pi1','.pi2','.pi3','.pict','.pct','.png','.pnm','.pns','.ppm','.psb','.pdd','.psp','.px','.pxm','.pxr','.qfx','.raw','.rle','.sct','.sgi','.tga','.tiff','.vtf','.xbm','.xcf','.xpm','.zif','.3dv','.amf','.awg','.ai','.cgm','.cdr','.cmx','.dp','.dxf','.e2d','.fs','.gbr','.odg','.svg','.stl','.vrml','.x3d','.sxd','.tgax','.v2d','.vdoc','.vsd','.vsdx','.vnd','.wmf','.emf','.xar'),
    #Program
    'webpage':('.dtd','.html','.xhtml','.mhtml','.maf','.maff','.asp','.aspx','.adp','.bml','.cfm','.cgi','.ihtml','.jsp','.las','.lasso','.lassoapp','.pl','.php','.ssi'),
    'program':('.adb','.ads','.ahk','.applescript','.as','.au3','.bat','.bas','.btm','.class','.cljs','.cmd','.coffee','.c','.cpp','.cs','.ino','.egg','.erb','.go','.hta','.ibi','.ici','.ijs','.ipynyb','.itcl','.js','.jsfl','.kt','.lua','.m','.mrc','.ncf','.nuc','.nud','.nut','.o','.pde','.pm','.ps1','.ps1xml','.psc1','.psd1','.psm1','.py','.pyc','.pyo','.pyw','.r','.rb','.rdp','.red','.rs','.sb2','.scpt','.scptd','.sdl','.sh','.syjs','.sypy','.tcl','.tns','.vbs','.xpl','.ebuild'),
    'database':('.4db','.4dd','.4dindy','.4dindx','.4dr','.accdb','.accde','.adt','.apr','.box','.chml','.daf','.dat','.db','.dbf','.dta','.ess','.eap','.fdb','.fp,','.frm','.gdb','.gtable','.kexi','.kexic','.kexis','.ldb','.lirs','.mda','.mdb','.adp','.mde','.mdf','.myd','.myi','.ncf','.nsf','.ntf','.nv2','.odb','.ora','.pcontact','.pdb','.pdi','.pdx','.prc','.sql','.rec','.rel','.rin','.sdb','.sdf','.sqlite','.udl','.wadata','.waindx','.wamodel','.wajournal','.wdb','.wmdb'),
    # File
    'comp_archive':('.?mn','.main','..?q?','.7z','.aapkg','.ace','.alz','.appx','.at3','.bke','.arc','.arj','.ass','.b','.ba','.big','.bjsn','.bkf','.bzip2','.bld','.cab','.c4','.cals','.xaml','.clipflair','.cpt,','.daa','.deb','.dmg','.ddz','.dn','.dpe','.ecab','.esd','.ess','.flipchart','.gbp','.gbs','.gho','.gzip','.ipg','.lbr','.lqr','.lha','.lzip','.lzo','.lzma','.lzx','.mbw','.mpq','.bin','.nl2pkg','.nth','.oar','.osk','.osr','.osz','.pak','.par','.paf','.pea','.rar','.rag,','.rax','.rbxl','.rbxlx','.rpm','.sb','.sb3','.sen','.sit','.sis','.sisx','.skb','.sq','.swm','.szs','.tar','.tgz','.tb','.tib','.uha','.uue','.vol','.vsa','.wax','.wim','.xap','.xz','.z','.zoo','.zip'),
    'executables':('.msi','.apk','.com','.exe','.gadget','.jar','.wsf'),
    'font':('.abf','.afm','.bdf','.bmf','.brfnt','.fnt','.fon','.mgf','.otf','.postscript','.pfa','.pfb','.pfm','.fond','.sfd','.snf','.tdf','.tfm','.ttf','.ufo','.woff'),
    'system':('.bak','.cfg','.cpl','.cur','.dll','.dmp','.drv','.icns','.ico','.ini','.sys'),
    'document':('.0','.1st','.600','.602','.abw','.acl','.afp','.ami','.amigaguide','.ans','.asc','.aww','.ccf','.csv','.cwk','.dbk','.dita','.doc','.docm','.docx','.dot','.dotm','.dotx','.epub','.ezw','.fdx','.ftm','.fx','.gdoc','.hwp','.hwpml','.log','.lwp','.mbp','.md','.me','.mcw','.mobi','.nb','.nbp','.neis','.nt','.nq','.odm','.odoc','.odt','.ods','.osheet','.ott','.omm','.pages','.pap','.pdax','.pdf','.quox','.radix-64','.rtf','.rpt','.sdw','.se','.stw','.sxw','.tex','.info','.troff','.txt','.uof','.uoml','.via','.wpd','.wps','.wpt','.wrd','.wrf','.xps','.pot','.potm','.potx','.ppa','.ppam','.pps','.ppsm','.ppsx','.ppt','.pptm','.pptx')
}

# function that looks for files in \\INPUT and moves them into their sorted folders
def InputToSorted(startDir, endDir):
    # get list of files in startDir
    files = os.listdir(startDir)

    # check if the list of files is not empty
    if len(files) != 0:
        # iterate code block for every file in files
        for f in files:
            # seperate the file name and file extension into seperate variables
            file_name, file_ext = os.path.splitext(f)

            # get the start path of the file
            fileStartPath = startDir + '\\' + file_name + file_ext
            
            # get the file type mapped to the file extention
            for i in file_ext_file_type:
                for j in file_ext_file_type[i]:
                    if j == file_ext:
                        file_ext_Key = i
            
            # get the folder path mapped to the file type
            file_type_folder_path = file_type_folder[file_ext_Key]
            # add the file type folder to the file path
            fileEndPath = endDir + '\\' + file_type_folder_path

            # add year to path
            year = date.today().year
            fileEndPath = fileEndPath + '\\' + str(year)
            #add month to path
            month = '{:02d}'.format(date.today().month)
            fileEndPath = fileEndPath + '\\' + str(month)

            # get the folder end path
            folderEndPath = fileEndPath

            # get the default end path of the file
            checkfileEndPath = fileEndPath + '\\' + file_name + file_ext
            
            # increament counter
            i = 0
            # check if the file exists already
            while os.path.exists(fileEndPath):
                # increate the iteration of duplicate file
                i+=1
                # create new file end path with the duplicate number
                checkfileEndPath = fileEndPath + '\\' + file_name + '_' + str(i) + file_ext
                if not os.path.exists(checkfileEndPath):
                    break
            if i != 0:
                fileEndPath = fileEndPath + '\\' + file_name + '_' + str(i) + file_ext
            else:
                fileEndPath = fileEndPath + '\\' + file_name + file_ext
            
            # create the subfolders if they dont exist
            if not os.path.exists(folderEndPath):
                os.makedirs(str(folderEndPath))

            # move file from its start path to its end path
            shutil.move(fileStartPath, fileEndPath)
            CREATE_NO_WINDOW = 0x08000000
            subprocess.call(r'C:\\Users\\Owner\\Documents\\RefreshDesktop\\refresh.bat', creationflags=CREATE_NO_WINDOW)

def DownloadsToDesktop(startDir, endDir):
    # get list of files in startDir
    files = os.listdir(startDir)

    # check if the list of files is not empty
    if len(files) != 0:
        # iterate code block for every file in files
        for f in files:
            # seperate the file name and file extension into seperate variable
            file_name, file_ext = os.path.splitext(f)

            if file_ext != '.tmp' and file_ext != '.crdownload':
                # get the start path of the file
                fileStartPath = startDir + '\\' + file_name + file_ext
                
                # get the default end path of the file
                fileEndPath = endDir + '\\' + file_name + file_ext
                
                # increament counter
                i = 0
                # check if the file exists already
                while os.path.exists(fileEndPath):
                    # increate the iteration of duplicate file
                    i+=1
                    # create new file end path with the duplicate number
                    fileEndPath = endDir + '\\' + file_name + '_' + str(i) + file_ext

                # move file from its start path to its end path
                shutil.move(fileStartPath, fileEndPath)
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                si.wShowWindow = subprocess.SW_HIDE # default
                subprocess.call(r'C:\\Users\\Owner\\Documents\\RefreshDesktop\\refresh.bat', startupinfo=si)

while True:
    # try and execpt for if the process crashes or is stopped and will break the loop
    try:
        # wait one second before running the functions
        time.sleep(1)
        # calling functions
        DownloadsToDesktop(r'C:\\Users\\Owner\Downloads',
                           r'C:\\Users\\Owner\Desktop')
        InputToSorted('C:\\Users\\Owner\\Desktop\\INPUT',
                      'C:\\Users\\Owner\\Desktop')

    except KeyboardInterrupt:
        break