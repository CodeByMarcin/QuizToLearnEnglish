import pandas as pd
from gtts import gTTS
import os

def read_answer(text):
    tts = gTTS(text=text, lang='en', slow=True)
    tts.save("eng.mp3")
    os.system("start eng.mp3")

def quiz_from_excel(file_path, num_questions=50):
    df = pd.read_excel(file_path, header=None)

    correct_answers = 0
    questions_asked = 0

    while questions_asked < num_questions:
        row = df.sample().iloc[0]

        if row[3] >= 8:
            continue

        print(str(row[2]))

        user_input = input("Provide the answer: ")
        questions_asked += 1

        if str(user_input).strip().lower() == str(row[0]).strip().lower():
            print("Correct!")
            correct_answers += 1
            df.loc[df.index == row.name, 3] += 1
        else:
            correct_answer = str(row[1])
            print(f"Wrong answer. Correct answer is: {correct_answer}")
            read_answer(correct_answer)
            if row[3] > 5:
                df.loc[df.index == row.name, 3] -= 1

        df.to_excel(file_path, index=False, header=None)

    print(f"Number of correct answers: {correct_answers} from {num_questions}")

file_path = 'A-E.xlsx'
quiz_from_excel(file_path)
