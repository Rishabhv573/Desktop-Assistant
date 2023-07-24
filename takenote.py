import os
import datetime as dt

def func(note):
    directory = "Openai"
    if not os.path.exists(directory):
        os.makedirs(directory)
    timenow = dt.datetime.now().strftime("%H%M")
    date = dt.date.today()
    file_path = f"Openai/note-{date}-{timenow}.txt"
    with open(file_path, 'w') as file:
        now = dt.datetime.now().strftime("%H:%M:%S")
        file.write(now + "\n")
        file.write(note)
    file.close()