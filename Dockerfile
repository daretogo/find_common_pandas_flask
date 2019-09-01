FROM python:3.7-alpine 

RUN pip install pandas, pandas_usaddress, pdb, flask

CMD ["sleep", "infinity"]