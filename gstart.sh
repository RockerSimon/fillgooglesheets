wsgi.py
from app import app

if __name__ == '__main__':
    app.run()

application = app  # or rename your Flask application object if necessary

if __name__ == '__main__':
    application.run(host='0.0.0.0',port=8000)

#no wsgi
#nohup flask run --host=172.17.0.1 &

#wsgi
nohup gunicorn -b 0.0.0.0:8000 wsgi:application &
