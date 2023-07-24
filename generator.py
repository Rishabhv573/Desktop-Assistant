import openai
from config import apikey_openai
import os
import datetime as dt

def func(prompt):
    openai.api_key = apikey_openai
    text = f"{prompt} \n*************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.makedirs("Openai")
    date = dt.date.today()
    with open(f"Openai/AI-generated-{date}.txt", "w") as f:
        f.write(text)
        f.close()