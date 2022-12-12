FROM python:3.7
COPY . /app
WORKDIR /app
RUN pip install -r Requirements.txt
RUN pip uninstall -y protobuf
RUN pip install protobuf==3.19.2
EXPOSE 8502
ENTRYPOINT ["streamlit","run"]
CMD ["Redaction_Contrat.py"]