FROM python:3-alpine

WORKDIR /backend

COPY requirements.txt /backend/
COPY . /backend/

RUN pip install --no-cache-dir -r requirements.txt

CMD ["python", "script.py"]
