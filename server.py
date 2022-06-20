from flask import Flask
from flask import render_template
from flask import Response, request, jsonify
app = Flask(__name__)

db = 'example.db'

# ROUTES
@app.route('/')
def welcome():
   return render_template('welcome.html')   


@app.route('/infinity')
def log_sales():
    return render_template('log_sales.html', clients = clients, sales = sales)  

# AJAX FUNCTIONS
# ajax for log_sales.js

@app.route('/save_sale', methods = ['GET', 'POST'])
def save_sale():
    global current_id
    global sales
    global clients
    
    new_sale_entry = request.get_json()   
    client = new_sale_entry["client"]

    current_id += 1
    new_sale_entry["id"] = current_id

    sales.insert(0, new_sale_entry)

    if client not in clients:
        clients.append(client)

    return jsonify(sales = sales, clients = clients)


@app.route('/delete_sale', methods = ['GET', 'POST'])
def delete_sale():
    global sales

    id = request.get_json()

    for sale in sales:
        if sale["id"] == id:
            sales.remove(sale)

    return jsonify(sales = sales)


if __name__ == '__main__':
   app.run(debug = True)