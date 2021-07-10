import PySimpleGUI as sg
import time
import configparser
import sys
import os
import webbrowser
import csv
from bs4 import BeautifulSoup
from lxml import html
import requests
import re
import logging
import inspect

# pyinstaller moodleprogram.py
# moodinstaller

abs_path = ''
for letter in os.path.dirname(os.path.abspath(__file__)) + '\\':
    abs_path += '\\' if letter == '\\' else ''
    abs_path += letter
logging.basicConfig(filename=abs_path + 'log\\all.log', format='%(asctime)s: %(levelname)-7s: %(message)s',
                    datefmt='%m/%d/%Y %H:%M:%S %z', level=logging.NOTSET)
chrome_path = abs_path + 'chrome\\'
abs_path += 'config\\'
clg_start_time = 10
line_length = 15
acro = {}
arg = sys.argv
gui = False
show_browser = False
auto = False
fieldnames = ['date', 'time', 'user', 'sub']


def read_acro():
    global acro
    with open(abs_path + "acro.csv", "r") as f:
        reader = csv.DictReader(f)
        for i in reader:
            acro[i['acro']] = i['full']
    return


def save_marking_stats(user):
    global fieldnames
    with open(abs_path + "marking_stats.csv", 'a', newline='') as f:
        csv_writer = csv.DictWriter(f, fieldnames=fieldnames)
        f.seek(0, 2)
        if f.tell() == 0:
            csv_writer.writeheader()
        csv_writer.writerow({"date": time.strftime('%d-%m'), "time": time.strftime('%H:%M'), 'user': user,
                             'sub': fetch_time_table(*fetch_time())})


def read_marking_stats():
    try:
        with open(abs_path + "marking_stats.csv", 'r') as f:
            csv_reader = csv.reader(f)
            data = list(csv_reader)
            if len(data) == 1 or len(data[1]) == 0:
                try:
                    data[1] = ['NO', 'DATA', 'IN', 'HERE']
                except IndexError:
                    data.append(['NO', 'DATA', 'IN', 'HERE'])
    except FileNotFoundError:
        raise_ex("Never performed auto atendance!")
        return ['NO', 'DATA', 'IN', 'HERE']
    except Exception as e:
        raise_ex(e, True)
        return
    else:
        return data[:0:-1]


def raise_ex(text, unknown=False, title=''):
    type_er = str(type(text).__name__)
    if type_er != 'str':
        text = type_er + ": " + str(text)
    if unknown:
        logging.error(f"{inspect.stack()[1][3]:<20}:{title} {str(text).strip()}")
    else:
        logging.warning(f"{inspect.stack()[1][3]:<20}:{title} {str(text).strip()}")
    if gui:
        sg.popup_error(text, title="Error " + title, non_blocking=True)
    else:
        raise Exception(f"{text}Open help or type 'help' or '-h' for help.")


def print_gui(text, title='', non_blocking=False):
    if gui:
        sg.popup(text, line_width=400, title=title, non_blocking=non_blocking)
    else:
        if title != '':
            print(f"{title}: ", end='')
        print(text)
    new_line = "\n"
    logging.info(f"{inspect.stack()[1][3]:<20}:{title} {str(text).strip().split(new_line)[0]}")


def is_marked(status, user_number=0):
    try:
        with open(abs_path + "isMarked.txt", "r") as f:
            stats = [i.split('-') for i in f.read().strip().split(':')]
    except FileNotFoundError:
        stats = [[str(i), '0', '0'] for i in range(total_user())]
        with open(abs_path + "isMarked.txt", "w") as f:
            f.write(':'.join(['-'.join(i) for i in stats]))
    except Exception as e:
        raise_ex(e, True)
        return

    for i in range(0, len(stats)):
        if len(stats[i]) != 3 or not all(stats[i]):
            stats[i] = [str(i), '0', '0']
    for i in range(len(stats), total_user()):
        stats.append([str(i), '0', '0'])

    if status:
        line = ''
        save_marking_stats(user_number)
        stats[user_number][1] = str(time.localtime().tm_mday)
        stats[user_number][2] = str(time.localtime().tm_hour)
        with open(abs_path + "isMarked.txt", "w") as f:
            for i in stats:
                line += '-'.join(i)
                line += ':'
            f.write(line[:-1])
        return

    user = []
    for i in range(total_user()):
        try:
            if int(stats[i][1]) != time.localtime().tm_mday or int(stats[i][2]) != time.localtime().tm_hour:
                user.append(i)
        except IndexError:
            user.append(i)
    return user


def do_browser(status):
    global show_browser
    if status:
        with open(abs_path + "doBrowser.txt", "w") as f:
            f.write("1" if show_browser else "0")
    try:
        with open(abs_path + "doBrowser.txt", "r") as f:
            show_browser = True if int(f.read().strip()) == 1 else False
    except FileNotFoundError:
        do_browser(True)
    except Exception as e:
        raise_ex(e, True)
    finally:
        return


def read_config(name):
    if os.path.exists(abs_path + name + '.ini'):
        config = configparser.ConfigParser()
        config.read(abs_path + name + ".ini")
        return config
    else:
        raise_ex("File not found! Either run the installer again or change directory.\n")
        return None


def save_config(name, dic):
    if os.path.exists(abs_path + name + '.ini'):
        config = configparser.ConfigParser()
        for i in dic.keys():
            config[i] = dic[i]
        with open(abs_path + name + '.ini', 'w') as configfile:
            config.write(configfile)
    else:
        raise_ex("File not found! Either run the installer again or change directory.\n")
        return None


def total_user():
    user_data = read_config('user')
    if user_data is None:
        return
    return len(user_data.sections()) - 1


def fetch_user(user_number):
    user_data = read_config('user')
    if user_data is None:
        return
    return {'user': user_data['login' + str(user_number)]['user'],
            'pass': user_data['login' + str(user_number)]['pass']}


def fetch_link(subject='', sta=False):
    user_data = read_config('user')
    if user_data is None:
        return
    if subject:
        try:
            return user_data['links'][subject]
        except KeyError:
            if not sta:
                raise_ex("Link not found!! \n")
        except Exception as e:
            raise_ex(e, True)
        return None


def read_whole_user():
    subjects = print_subject(1).keys()
    user = {}
    for i in range(total_user()):
        user['login' + str(i)] = fetch_user(i)
    user['links'] = {}
    for i in subjects:
        if i != '':
            j = fetch_link(i, True)
            if j is not None:
                user['links'][i] = j
    return user


def browser(user_number=0, sub=''):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    import selenium.common.exceptions as ex
    chrome_options = Options()
    chrome_options.add_experimental_option("useAutomationExtension", False)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", 'enable-logging'])
    chrome_options.add_argument("user-data-dir=" + chrome_path)
    if len(sub) == 0:
        chrome_options.add_experimental_option("detach", True)
        sub1 = list(print_subject(1).keys())
        for i in ['', 'NPTEL', 'LUNCH']:
            sub1.remove(i)
        sub1 = sub1[0]
        link = fetch_link(sub1)
    else:
        link = fetch_link(sub)

    if link is None:
        raise_ex("Link not found!!")
    try:
        driver = webdriver.Chrome(options=chrome_options, executable_path=abs_path + "chromedriver.exe")
    except Exception as e:
        raise_ex(e, True, f"User #{user_number}")
        if sub != '':
            print_gui("Trying without browser!", title=f'User #{user_number+1}', non_blocking=True)
            return cmd(user_number, sub)
        return
    driver.get(link)
    username = driver.find_element_by_name("username")
    password = driver.find_element_by_name("password")
    login_url = driver.current_url
    user_details = fetch_user(user_number)
    if user_details is None:
        return
    username.send_keys(user_details['user'])
    password.send_keys(user_details['pass'])
    driver.find_element_by_id("loginbtn").click()
    if driver.current_url == login_url:
        time.sleep(3)
        driver.quit()
        raise_ex("Wrong credentials. \n", title=f"User #{user_number}")
        return
    if sub:
        try:
            driver.find_element_by_link_text("Submit attendance").click()
            driver.find_element_by_name("status").click()
            driver.find_element_by_name("submitbutton").click()
            if auto:
                is_marked(True, user_number)
            driver.quit()
            return True
        except ex.NoSuchElementException:
            time.sleep(3)
            driver.quit()
            raise_ex("Option Not present!! Either Already Marked or Not Open for Self Marking.",
                     title=f"User #{user_number}")
            return
        except Exception as e:
            raise_ex(e, True, f"User #{user_number}")
            print_gui("Trying without browser!", title=f'User #{user_number+1}', non_blocking=True)
            return cmd(user_number, sub)
    else:
        driver.get(login_url.replace("login/index.php", "my/"))


def cmd(user_number, sub=''):
    link = fetch_link(sub)
    if link is None:
        return
    user_details = fetch_user(user_number)
    if user_details is None:
        return
    session = requests.session()
    result = session.get(link)
    login_url = result.url
    logintoken = list(set(html.fromstring(result.text).xpath("//input[@name='logintoken']/@value")))[0]
    username = user_details['user']
    password = user_details['pass']
    result = session.post(login_url, data=dict(username=username, password=password, logintoken=logintoken),
                          headers=dict(referer=login_url))

    if result.url == login_url:
        raise_ex("Wrong credentials. \n", title=f"User #{user_number}")
        return

    home_url = login_url[:-15]

    result = session.get(link, headers=dict(referer=link))
    soup = BeautifulSoup(result.content, 'html.parser')
    a = soup.select('a[href^="' + home_url + 'mod/attendance/attendance.php?"]')

    if len(a) == 0:
        raise_ex("Option Not present!! Either Already Marked or Not Open for Self Marking.",
                 title=f"User #{user_number}")
        return

    attend_link = (re.search('(http.*)"', str(a[0]))).group()[:-1]
    attend_page = session.get(attend_link, headers=dict(referer=link), allow_redirects=True)
    current_url = attend_page.url
    sesskey = current_url.split("=")[-1]
    sessid = current_url[current_url.index('=') + 1:current_url.index('&')]

    soup = BeautifulSoup(attend_page.content, 'lxml')
    lables = soup.find("div", attrs={"class": "d-flex flex-wrap align-items-center"})
    status = lables.find_all("input")[0]['value']

    result = session.post(attend_link,
                          data={'sessid': sessid, 'sesskey': sesskey, 'sesskey': sesskey,
                                '_qf__mod_attendance_student_attendance_form': 1,
                                'mform_isexpanded_id_session': 1, 'status': status,
                                'submitbutton': "Save+changes"},
                          headers=dict(referer=attend_link), allow_redirects=True)

    if result.url != attend_page.url:
        if auto:
            is_marked(True, user_number)
        return True
    else:
        return


def fetch_time_table(day="all", period="all", show_popup=True):
    time_table = read_config('timetable')
    if time_table is None:
        return
    days = time_table.sections() if day == "all" else [day]

    if days[0] not in time_table.sections():
        if show_popup:
            raise_ex("No Class Today!")
        return None
    if period != "all":
        if int(time_table[days[0]]['start']) > int(period):
            if show_popup:
                raise_ex(f"No Class, it will start from {time_table[days[0]]['start']}. \n")
            return None
        elif int(time_table[days[0]]['end']) <= int(period):
            if show_popup:
                raise_ex(f"No Class, college is over for today, time - {time_table[days[0]]['end']}. \n")
            return None
        return time_table[days[0]][period]
    else:
        line = ("{:<" + str(line_length) + "}" + (("{:^" + str(line_length) + "}") * 7)).format("Time", "10:00-11:00",
                                                                                                "11:00-12:00",
                                                                                                "12:00-13:00",
                                                                                                "13:00-14:00",
                                                                                                "14:00-15:00",
                                                                                                "15:00-16:00",
                                                                                                "16:00-17:00") + "\n"
        for i in days:
            flag = False
            time_table_day = time_table[i]
            line += f"{i:<{line_length}}"
            if int(time_table_day['start']) > clg_start_time:
                line += " " * line_length * (int(time_table_day['start']) - clg_start_time)

            for j in range(int(time_table_day['start']), int(time_table_day['end'])):
                if flag:
                    flag = False
                elif j != int(time_table_day['end']) - 1 and time_table_day[str(j)] == time_table_day[str(j + 1)]:
                    line += f"{time_table_day[str(j)]:^{line_length * 2}}"
                    flag = True
                else:
                    line += f"{time_table_day[str(j)]:^{line_length}}"
            line += '\n'
        print_gui(line)


def fetch_time():
    return (time.strftime("%A", time.localtime())), (time.strftime("%H", time.localtime()))


def mark_attendance(users, subject=''):
    if len(subject) == 0:
        day, hour = fetch_time()
        subject = fetch_time_table(day, hour)
        if subject is None:
            return
        elif subject == '':
            raise_ex("Its Free period!! \n")
            return
        elif subject == "LUNCH":
            raise_ex("Its LUNCH, eat food!! \n")
            return
        print_line = acro[subject] if subject in acro.keys() else subject
        print_line += "\nDAY: " + day + "  |  TIME: " + hour
        print_gui(print_line, title="Period", non_blocking=True)
    for user in users:
        try:
            if (show_browser and browser(user, subject)) or (not show_browser and cmd(user, subject)):
                print_gui("Attendance marked!", non_blocking=True, title=f'User #{int(user)+1}')
        except Exception as e:
            print_gui(f"For User {int(user) + 1}\n{e}")


def print_subject(return_list=0):
    subject = read_config('timetable')
    if subject is None:
        return
    day = subject.sections()
    sub = {}
    for i in day:
        for j in range(int(subject[i]['start']), int(subject[i]['end'])):
            sub[subject[i][str(j)]] = sub.setdefault(subject[i][str(j)], 0) + 1
    if return_list:
        return sub
    print_line = "{:>20}\t{:<5}\n".format('Subject', 'Number of hours of lecture')
    for i in sub.keys():
        if i == "LUNCH":
            continue
        if len(i) == 0:
            continue
        print_line += f"{i:>20}\t{sub[i]:<5}\n"
    print_line += ("-" * 60) + "\n"
    print_line += ("{:>20}\t{:<20}\n".format("Acronym", 'Subject'))
    for i in acro.keys():
        print_line += f"{i:>20}\t{acro[i]:<20}\n"
    print_gui(print_line)


def change_login():
    dic = read_whole_user()
    while True:
        print_gui('''\n1. If you want to add user, write 'username password'
2. If you want to delete the last added user write 'delete'
3. If you want to change the user details write 'change'
4. To exit press 'ENTER'\n''')
        new_cmd = input()
        no_of_users = len(dic) - 1
        if new_cmd == 'delete':
            if no_of_users > 1:
                dic.pop('login' + str(no_of_users - 1))
            else:
                print_gui("1 User already, cannot delete all users!")
        elif new_cmd == 'change':
            print_gui("If you want to change write the value, else press \"ENTER\"")
            print_gui(f"Total number of Users:{no_of_users}")
            for i in range(no_of_users):
                print_gui(f"For user {i + 1}")
                print_gui(f"Current Username: {dic['login' + str(i)]['user']}")
                new_credential = input()
                if new_credential != '' and new_credential is not None:
                    dic['login' + str(i)]['user'] = new_credential
                print_gui(f"Current Password: {dic['login' + str(i)]['pass']}")
                new_credential = input()
                if new_credential != '' and new_credential is not None:
                    dic['login' + str(i)]['pass'] = new_credential
        elif len(new_cmd.split(' ')) > 1:
            detail = new_cmd.split(' ')
            dic['login' + str(no_of_users)] = {}
            dic['login' + str(no_of_users)]['user'] = detail[0]
            dic['login' + str(no_of_users)]['pass'] = ' '.join(detail[1:])
        elif new_cmd == '':
            break
    save_config('user', dic)
    print_gui("Changes successful!")
    return


def change_links():
    dic = read_whole_user()
    sub = []
    for i in print_subject(1).keys():
        if i not in ['', 'LUNCH', 'NPTEL']:
            sub.append(i)
    print_gui("If you want to change write the value, else press \"ENTER\" or to delete value write 'delete'")
    lin = {}
    i = 0
    while True:
        print_gui(
            "Current link for {}: {}".format(sub[i], dic['links'][sub[i]] if sub[i] in dic['links'].keys() else ''))
        lin[sub[i]] = input()
        if lin[sub[i]] is not None and lin[sub[i]] != '' and not lin[sub[i]].startswith("https://") and \
                not lin[sub[i]].startswith("http://") and lin[sub[i]].lower() != 'delete':
            print_gui("Enter https:// or http:// URL only!")
            i -= 1
        i += 1
        if i == len(sub):
            break

    for i in lin.keys():
        if lin[i] is not None and lin[i] != '':
            if lin[i].lower() == 'delete':
                dic['links'].pop(i)
            else:
                dic['links'][i] = lin[i]
    save_config('user', dic)
    print_gui("Changes successful!")


def print_marking_stats():
    data = read_marking_stats()
    line = "\n\t{:<6} {:<6} {:<5} {:<20}\n".format(*[i.upper() for i in fieldnames])
    for i in data:
        line += "\t{:<6} {:<6} {:<5} {:<20}\n".format(*i)
    print_gui(line[:-1])
    return


def print_version():
    print_gui("Version 2.0, Developed by Mysterious-Owl\nhttps://mysteriousowl.tech/\n"
              "To check for updates go to-\nhttps://github.com/Mysterious-Owl/moodle/releases/latest")


def print_help():
    print_gui('''
This program is basically to mark attendance and open moodle.    
Commands:
  (no arguments)\t\tOpen GUI console.
  auto\t\t\t\tWill mark attendance according to current time.
  mark <sub>\t\t\tMark attendance of this particular subject.
  open\t\t\t\tOpens moodle.
  tt\t\t\t\tPrint time table.
  subject\t\t\tPrints the list of all subject.
  print\t\t\t\tPrints the auto attendance marking stats.
  change <>\t\t\tChange <login> or <links>
  help\t\t\t\tShow help.

General Options:
  -h, --help\t\t\tShow help.
  -v, --version\t\t\tShow version.
  -s, --sub\t\t\tPrints the list of all subject.
  -a\t\t\t\tWill mark attendance according to current time.
  -m <sub>\t\t\tMark attendance of this particular subject.
  -o\t\t\t\tOpens moodle.
  -c <>\t\t\t\tChange <login> or <links>
  -p\t\t\t\tPrints the auto attendance marking stats.

Flags:
  --b\t\t\t\tFlag, can be used when marking attendance to open browser to mark attendance.
''', "CLI Help")


def print_help_gui():
    layout = [[sg.Text("This program is basically to mark attendance and open moodle.", font=("", 12))],
              [sg.Text("")],
              [sg.Text("Automate attendance:", font=("", 12))],
              [sg.Text("1. To automate attendance goto:")],
              [sg.Text("Start -> Windows Administrative Tools -> Task Scheduler or search \"Task Scheduler\"")],
              [sg.Text(
                  "2. Click on \"Task Scheduler Library\" on right side then on left side click on \"Import Task\".")],
              [sg.Text("3. Browse to \"" + abs_path[:-22].replace("\\\\", "\\") +
                       "\" and select \"moodle.xml\" and then OK then OK.")],
              [sg.Text("4. Done")],
              [sg.Text("")],
              [sg.Text("Command Line Arguments:", font=("", 12))],
              [sg.Text("Commands :", font=("", 11))],
              [sg.Text("  "), sg.Text("no arguments", size=(14, 1)), sg.Text("Open GUI console.")],
              [sg.Text("  "), sg.Text("auto", size=(14, 1)), sg.Text("Will mark attendance according to current time")],
              [sg.Text("  "), sg.Text("mark <sub>", size=(14, 1)),
               sg.Text("Mark attendance of this particular subject.")],
              [sg.Text("  "), sg.Text("open", size=(14, 1)), sg.Text("Opens moodle.")],
              [sg.Text("  "), sg.Text("tt", size=(14, 1)), sg.Text("Print time table.")],
              [sg.Text("  "), sg.Text("subject", size=(14, 1)), sg.Text("Prints the list of all subject.")],
              [sg.Text("  "), sg.Text("print", size=(14, 1)), sg.Text("Prints the auto attendance marking stats.")],
              [sg.Text("  "), sg.Text("change <>", size=(14, 1)), sg.Text("Change <login> or <links>")],
              [sg.Text("  "), sg.Text("help", size=(14, 1)), sg.Text("Show help.")],
              [sg.Text("")],
              [sg.Text("General Options:", font=("", 11))],
              [sg.Text("  "), sg.Text("-h, --help", size=(14, 1)), sg.Text("Show help.")],
              [sg.Text("  "), sg.Text("-v, --version", size=(14, 1)), sg.Text("Show version.")],
              [sg.Text("  "), sg.Text("-s, --sub", size=(14, 1)), sg.Text("Prints the list of all subject.")],
              [sg.Text("  "), sg.Text("-a", size=(14, 1)), sg.Text("Will mark attendance according to current time.")],
              [sg.Text("  "), sg.Text("-m <sub>", size=(14, 1)),
               sg.Text("Mark attendance of this particular subject.")],
              [sg.Text("  "), sg.Text("-o", size=(14, 1)), sg.Text("Opens moodle.")],
              [sg.Text("  "), sg.Text("-c <>", size=(14, 1)), sg.Text("Change <login> or <links>")],
              [sg.Text("  "), sg.Text("-p", size=(14, 1)), sg.Text("Prints the auto attendance marking stats")],
              [sg.Text("")],
              [sg.Text("Flags:", font=("", 11))],
              [sg.Text("  "), sg.Text("--b", size=(14, 1)), sg.Text("Flag, can be used when marking attendance to open"
                                                                    " browser to mark attendance.")],
              [sg.Text("  ")],
              [sg.Text('')], [sg.Button('Back', size=(5, 1), key="Exit"), sg.Text("", size=(33, 1)),
                              sg.Button('More Help/Report Issue', key='issue')]
              ]

    window_help = sg.Window("Help", layout, default_element_size=(48, 2), text_justification='l',
                            auto_size_text=True, auto_size_buttons=False, default_button_element_size=(24, 1),
                            finalize=True)
    while 1:
        event, value = window_help.Read()
        if event == 'issue':
            webbrowser.open("https://github.com/Mysterious-Owl/moodle")
        elif event == 'Exit':
            window_help.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_help.Close()
            sys.exit(0)


def select_subject_gui():
    sub = [i for i in print_subject(1).keys() if i not in ['', 'NPTEL', 'LUNCH']]
    sub1 = [acro.get(n, n) for n in sub]

    layout = [[sg.Button(f'{j}', key=i)] for i, j in zip(sub, sub1)]
    layout += [sg.Text('')], [sg.Button('Back', size=(5, 1), key="Exit")]
    window_sub = sg.Window("Select Subject", layout, default_element_size=(48, 2), text_justification='c',
                           auto_size_text=True, auto_size_buttons=False, default_button_element_size=(48, 2),
                           finalize=True)
    while 1:
        event, value = window_sub.Read()
        if event in sub:
            mark_attendance([i for i in range(total_user())], event)
        elif event == 'Exit':
            window_sub.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_sub.Close()
            sys.exit(0)


def print_version_gui():
    layout = [[sg.Text('Version 2.0, Developed by Mysterious-Owl', size=(35, 1), font=("agency fb", 15))],
              [sg.Text("", size=(1, 1))],
              [sg.Button('Open Website', key='web')],
              [sg.Button('Check Update', key='upd')],
              [sg.Text('')],
              [sg.Button('Back', size=(5, 1), key="Exit")]
              ]
    window_ver = sg.Window("About", layout, default_element_size=(48, 2), text_justification='c',
                           auto_size_text=True, auto_size_buttons=False, default_button_element_size=(30, 2),
                           finalize=True, keep_on_top=True)
    while 1:
        event, value = window_ver.Read()
        if event == 'web':
            webbrowser.open('https://mysterious-owl.github.io/')
        elif event == 'upd':
            webbrowser.open('https://github.com/Mysterious-Owl/moodle/releases/latest')
        elif event == 'Exit':
            window_ver.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_ver.Close()
            sys.exit(0)


def print_time_table_gui():
    time_table = read_config('timetable')
    if time_table is None:
        return
    days = time_table.sections()
    timet = [["Time", "10:00-11:00", "11:00-12:00", "12:00-13:00", "13:00-14:00", "14:00-15:00", "15:00-16:00",
              "16:00-17:00"]]
    for i in days:
        flag = False
        time_table_day = time_table[i]
        tt = [i]
        if int(time_table_day['start']) > clg_start_time:
            tt.extend([""] * (int(time_table_day['start']) - clg_start_time))
        for j in range(int(time_table_day['start']), int(time_table_day['end'])):
            if flag:
                flag = False
            elif j != int(time_table_day['end']) - 1 and time_table_day[str(j)] == time_table_day[str(j + 1)]:
                tt.append(time_table_day[str(j)] + 'double')
                flag = True
            else:
                tt.append(time_table_day[str(j)])
        timet.append(tt)

    day = time.strftime('%A')
    layout = []
    exclusion = ['Time', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    for i in timet:
        lay = []
        for t in i:
            l = 30 if t.endswith("double") else 15
            t = t[:-6] if t.endswith("double") else t
            if t in exclusion:
                lay += sg.Text(t, size=(l, 1), justification='l', text_color="yellow", font=("", 12)),
            elif i[0] == day and fetch_time_table(*fetch_time(), show_popup=False) == t:
                lay += sg.Text(t, size=(l, 1), justification='c', font=("", 12), relief='ridge'),
            else:
                lay += sg.Text(t, size=(l, 1), justification='c', font=("", 12)),
        layout += [lay]
    layout += [sg.Text('')], [sg.Button('Back', size=(5, 1), key="Exit")]
    window_tab = sg.Window("Time Table", layout, default_element_size=(10, 2), text_justification='c',
                           auto_size_text=True, auto_size_buttons=False, finalize=True, keep_on_top=True)
    while 1:
        event, value = window_tab.Read()
        if event == 'Exit':
            window_tab.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_tab.Close()
            sys.exit(0)


def print_subject_gui():
    sub = print_subject(1)
    for i in ["", "LUNCH"]:
        if i in sub.keys():
            sub.pop(i)
    layout = []
    lay = sg.Text("Subject", size=(14, 1), justification='r', font=("", 16), text_color='sky blue'), sg.Text("  "), \
        sg.Text("Number of hours of lecture", size=(22, 1), justification='l', text_color='sky blue', font=("", 16))
    layout += [lay]
    for i in sub.keys():
        lay = sg.Text(i, size=(18, 1), justification='r', font=("", 12)), sg.Text("    "), \
              sg.Text(sub[i], size=(18, 1), justification='l', font=("", 12))
        layout += [lay]
    lay = sg.Text("-" * 70, size=(28, 1), font=("", 20), text_color='red'),
    layout += [lay]

    lay = sg.Text("Acronym", size=(14, 1), justification='r', font=("", 16), text_color='sky blue'), sg.Text("  "), \
        sg.Text("Full form", size=(22, 1), justification='l', text_color='sky blue', font=("", 16))
    layout += [lay]
    for i in acro.keys():
        lay = sg.Text(i, size=(18, 1), justification='r', font=("", 12)), sg.Text("    "), \
              sg.Text(acro[i], size=(27, 1), justification='l', font=("", 12))
        layout += [lay]
    layout += [sg.Text('')], [sg.Button('Back', size=(5, 1), key="Exit")]
    window_tab = sg.Window("Subject List", layout, default_element_size=(10, 2), auto_size_text=True,
                           auto_size_buttons=False, finalize=True, keep_on_top=True)
    while 1:
        event, value = window_tab.Read()
        if event == 'Exit':
            window_tab.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_tab.Close()
            sys.exit(0)


def change_login_gui():
    global gui
    gui = True
    dic = read_whole_user()
    current_no = total_user()
    lay = []
    already_preset = []
    for i in range(current_no):
        lay_temp = [[sg.Text('Username:', size=(8, 1)), sg.Input(dic['login' + str(i)]['user'], key=f'user{i}')],
                    [sg.Text('Password:', size=(8, 1)), sg.Input(dic['login' + str(i)]['pass'], key=f'pass{i}')]]
        lay += [[sg.Frame(f'User #{i + 1}', lay_temp, key=f'detail{i}')], [sg.Text('')]]
    layout = [[sg.Text('Enter login details -', size=(20, 2), text_color='yellow')],
              [sg.Column(layout=lay, key="details")],
              [sg.Text('', size=(10, 1)), sg.Button('ADD', key='add', size=(6, 1)), sg.Text('', size=(10, 1)),
               sg.Button('DELETE', key='delete', disabled=True if current_no == 1 else False)],
              [sg.Text('')],
              [sg.Text('', size=(35, 1)), sg.Button('Cancel', key='exit'), sg.Button('Submit', key='submit')]]
    login_change = sg.Window('Settings', layout, finalize=True)
    while 1:
        events, value = login_change.Read()
        if events == 'submit':
            a = min(total_user(), current_no)
            for i in range(a):
                dic['login' + str(i)]['user'] = value['user' + str(i)]
                dic['login' + str(i)]['pass'] = value['pass' + str(i)]
            for i in range(a, max(current_no, total_user())):
                if total_user() > current_no:
                    dic.pop('login' + str(i))
                else:
                    dic['login' + str(i)] = {}
                    dic['login' + str(i)]['user'] = value['user' + str(i)]
                    dic['login' + str(i)]['pass'] = value['pass' + str(i)]
            save_config('user', dic)
            login_change.Close()
            break
        elif events == 'add':
            if value[f'user{current_no-1}'].strip() == '' or value[f'pass{current_no-1}'].strip() == '':
                raise_ex(f"Enter value for User #{current_no} first")
                continue
            if current_no in already_preset:
                login_change[f'detail{current_no}'].update(visible=True)
                login_change[f'user{current_no}'].update('')
                login_change[f'pass{current_no}'].update('')
                already_preset.remove(current_no)
            else:
                lay_temp = [[sg.Text('Username:', size=(8, 1)), sg.Input(key=f'user{current_no}')],
                            [sg.Text('Password:', size=(8, 1)), sg.Input(key=f'pass{current_no}')]]
                login_change.extend_layout(login_change['details'],
                                           [[sg.Frame(f'User #{current_no + 1}', lay_temp, key=f'detail{current_no}')],
                                            [sg.Text('')]])
            current_no += 1
            if current_no > 1:
                login_change['delete'].update(disabled=False)
        elif events == 'delete':
            current_no -= 1
            login_change[f'detail{current_no}'].update(visible=False)
            already_preset.append(current_no)
            if current_no == 1:
                login_change['delete'].update(disabled=True)
        elif events == 'exit':
            login_change.Close()
            break
        elif events == sg.WIN_CLOSED:
            login_change.Close()
            sys.exit(0)


def change_link_gui():
    dic = read_whole_user()
    sub = []
    lin = []
    for i in print_subject(1).keys():
        if i not in ['', 'LUNCH', 'NPTEL']:
            sub.append(i)
            if i in dic['links'].keys():
                lin.append(dic['links'][i])
            else:
                lin.append('')
    layout = [[sg.Text(i, size=(15, 1), text_color='yellow'), sg.Input(j, (60, 1), key=i)] for i, j in zip(sub, lin)]
    layout += [sg.Text('')], [sg.Text('', size=(55, 1)), sg.Button('Cancel', key='Exit'),
                              sg.Button('Submit', key='submit')]
    window_link = sg.Window("Link Settings", layout, finalize=True)
    while 1:
        event, value = window_link.Read()
        flag = False
        if event == 'submit':
            for i in sub:
                if value[i] is None or value[i] == '':
                    if i in dic['links'].keys():
                        dic['links'].pop(i)
                    continue
                else:
                    if value[i].startswith('http://') or value[i].startswith('https://'):
                        dic['links'][i] = value[i]
                    else:
                        sg.popup_error("Not a valid address!!")
                        flag = True
            if flag:
                continue
            save_config('user', dic)
            window_link.Close()
            break
        if event == 'Exit':
            window_link.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_link.Close()
            sys.exit(0)


def setting_gui():
    global show_browser
    layout = [[sg.Button('Change LOGIN details', key='log'),
               sg.Button('Change LINKS', key='lin')],
              [sg.Checkbox('Show Browser while marking attendance?', key='show_b', default=show_browser,
                           enable_events=True)],
              [sg.Text('')],
              [sg.Button('Back', size=(5, 1), key="Exit")]]
    window_set = sg.Window("Settings", layout, default_element_size=(48, 2), text_justification='c',
                           auto_size_text=True, auto_size_buttons=False, default_button_element_size=(24, 2),
                           finalize=True, keep_on_top=True)
    while 1:
        event, value = window_set.Read()
        if event == 'log':
            window_set.Hide()
            change_login_gui()
            window_set.UnHide()
        elif event == 'lin':
            window_set.Hide()
            change_link_gui()
            window_set.UnHide()
        elif event == 'show_b':
            show_browser = value['show_b']
            do_browser(True)
        elif event == 'Exit':
            window_set.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_set.Close()
            sys.exit(0)


def marking_stats_gui():
    data = read_marking_stats()
    layout = [[sg.Table(data, headings=[i.upper() for i in fieldnames], hide_vertical_scroll=True, num_rows=10)],
              [sg.Text('')],
              [sg.Button('Back', size=(5, 1), key="Exit")]]
    window_stats = sg.Window("Marking Stats", layout)
    while 1:
        event, value = window_stats.read()
        if event == 'Exit':
            window_stats.Close()
            break
        elif event == sg.WIN_CLOSED:
            window_stats.Close()
            sys.exit(0)


def show_gui():
    global gui
    global show_browser
    global auto
    auto = False
    do_browser(False)
    gui = True
    os.system("mode con cols=16 lines=1")
    sg.theme('Daurk')
    sg.set_options(element_padding=(0, 0))

    layout = [[sg.Text('Select desired Option', size=(21, 1), font=("cooper black", 20))],
              [sg.Button('Open Moodle', key='-open-')],
              [sg.Button('Mark Automatically', key='-mark-')],
              [sg.Button('Mark for Specific Subject', key='-msub-')],
              [sg.Button('Show Time Table', key='-tt-')],
              [sg.Button('Show Subject list', key='-sub-')],
              [sg.Button('Show Auto-Marking stats', key='-amar-')],
              [sg.Button('Settings', key='-set-')],
              [sg.Button('About', key='-ver-')],
              [sg.Button('Help!!', key='-help-')]]

    window_main = sg.Window("Moodle", layout, default_element_size=(48, 2), text_justification='c',
                            auto_size_text=True, auto_size_buttons=False, default_button_element_size=(48, 2),
                            finalize=True)
    while True:
        event, values = window_main.read()
        if event == sg.WIN_CLOSED:
            break
        if event == '-open-':
            browser()
        elif event == '-mark-':
            auto = True
            choice = "Yes"
            users = is_marked(False)
            if len(users) == 0:
                choice = sg.popup_yes_no("As per local record attendance already marked."
                                         "\nWish to mark again?", title="Marked!!")
            if choice == "Yes":
                mark_attendance(users)
            auto = False
        elif event == '-msub-':
            window_main.Hide()
            select_subject_gui()
            window_main.UnHide()
        elif event == '-tt-':
            window_main.Hide()
            print_time_table_gui()
            window_main.UnHide()
        elif event == '-sub-':
            window_main.Hide()
            print_subject_gui()
            window_main.UnHide()
        elif event == '-amar-':
            window_main.Hide()
            marking_stats_gui()
            window_main.UnHide()
        elif event == '-set-':
            window_main.Hide()
            setting_gui()
            window_main.UnHide()
        elif event == '-ver-':
            window_main.Hide()
            print_version_gui()
            window_main.UnHide()
        elif event == '-help-':
            window_main.Hide()
            print_help_gui()
            window_main.UnHide()
    window_main.close()
    sys.exit(0)


read_acro()
if '--b' in arg:
    show_browser = True

if len(arg) == 1:
    show_gui()
elif arg[1] == "auto" or arg[1] == '-a':
    auto = True
    usersc = is_marked(False)
    if len(usersc) == 0:
        print_gui("Attendance already marked!")
    else:
        mark_attendance(usersc)
    auto = False
elif arg[1] == 'tt':
    fetch_time_table()
elif arg[1] == '-s' or arg[1] == 'subject' or arg[1] == '--sub':
    print_subject()
elif arg[1] == '-m' or arg[1] == 'mark':
    if len(arg) == 2 or arg[2].lower() not in [x.lower() for x in print_subject(1).keys()]:
        raise_ex("Enter correct subject! \n")
    else:
        mark_attendance([i for i in range(total_user())], arg[2])
elif arg[1] == '-o' or arg[1] == 'open':
    browser()
elif arg[1] == '-p' or arg[1] == 'print':
    print_marking_stats()
elif arg[1] == '-h' or arg[1] == '--help' or arg[1] == 'help':
    print_help()
elif arg[1] == '-v' or arg[1] == '--version':
    print_version()
elif arg[1] == '-c' or arg[1] == 'change':
    if len(arg) == 2 or arg[2] not in ['login', 'links']:
        raise_ex("Enter correct argument")
    elif arg[2] == 'login':
        change_login()
    elif arg[2] == 'links':
        change_links()
else:
    raise_ex(f"Unknown command \"{arg[1]}\" \n")
