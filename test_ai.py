import os
from openai import OpenAI

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

resp = client.responses.create(
    model="gpt-4.1-mini",
    input="Napisz dokładnie: dziala"
)

print(resp.output_text)
