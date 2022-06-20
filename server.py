from readline import append_history_file
from flask import Flask
from flask import render_template
from flask import Response, request, jsonify

from sqlalchemy.sql import text

app = Flask(__name__)

db_name = 'example.db'

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + db_name
app.config['SQLACHEMY_TRACK_MODIFICATIONS'] = True

db = SQLAlchemy(app)

# ROUTES
@app.route('/')
def welcome():
   return render_template('welcome.html')  

def testdb():
    try:
        db.session.query(text('1')).from_statement(text('SELECT 1')).all()
        return '<h1>It works.</h1>'
    except Exception as e:
        # e holds description of the error
        error_text = "<p>The error:<br>" + str(e) + "</p>"
        hed = '<h1>Something is broken.</h1>'
        return hed + error_text 


@app.route('/data')
def display():
    return render_template('display.html')  

# AJAX FUNCTIONS
# ajax for log_sales.js

@app.route('/add_emails', methods = ['GET', 'POST'])
def add_emails():
    
    return jsonify()


@app.route('/search', methods = ['GET', 'POST'])
def search():

    return jsonify()


if __name__ == '__main__':
   app.run(host='10.198.149.41', port=5000, debug=True, threaded=False)