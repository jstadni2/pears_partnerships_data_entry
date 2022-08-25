FROM python:3.9

WORKDIR /pears_partnerships_data_entry

COPY . .

RUN pip install -r requirements.txt

CMD [ "python", "./pears_partnerships_data_entry.py" ]