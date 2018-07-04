from subprocess import Popen, PIPE
import xlrd

def main():

    # AppleScript that opens up the iMessage app, creates a new text, and sends it
    apple_script = '''
        on run {targetBuddyPhone, targetMessage}
            tell application "System Events"
                tell application "Messages" to activate
                tell process "Messages"
                    click menu item "New Message" of menu "file" of menu bar 1
                    set input to targetMessage as text
                    delay 1
                    keystroke targetBuddyPhone
                    keystroke return
                    keystroke targetMessage
                    delay 3
                    keystroke return
                    delay 3
                end tell
            end tell
        end run
    '''

    text_msg = '''<Text Message Body>'''

    workbook = xlrd.open_workbook('<Excel Sheet Name>.xlsx')
    sheet = workbook.sheet_by_index(0)

    print('Starting to send text messages...')

    count = 1
    for cell in sheet.col(0):
        phone_number = str(cell.value)
        print('Line {}: Sending text to phone number {}'.format(count, phone_number))

        p = Popen(['osascript', '-'] + [phone_number, text_msg], stdin=PIPE, stdout=PIPE, stderr=PIPE, universal_newlines=True)
        stdout, stderr = p.communicate(apple_script)

        print('Finished sending text. Return code {}, printout {}, stderr {}'.format(p.returncode, stdout, stderr))
        count += 1

if __name__ == "__main__":
    main()