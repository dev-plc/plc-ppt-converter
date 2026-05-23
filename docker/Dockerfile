FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY convert_pptx.py canva_api.py telegram_bot.py ./
COPY cover.png ./

CMD ["python", "telegram_bot.py"]
