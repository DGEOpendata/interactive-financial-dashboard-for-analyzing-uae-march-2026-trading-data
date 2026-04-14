python
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from flask import Flask, render_template, request

# Create a Flask app
app = Flask(__name__)

# Load the dataset
data = pd.read_excel('Trading_Summary_March_2026.xlsx')

# Preprocess the data
data['Date'] = pd.to_datetime(data['Date'])
data.set_index('Date', inplace=True)

# Define a function to filter data
def filter_data(start_date, end_date):
    return data.loc[start_date:end_date]

# Define routes for the Flask app
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    filtered_data = filter_data(start_date, end_date)

    # Generate visualizations
    plt.figure(figsize=(10, 6))
    sns.lineplot(data=filtered_data, x=filtered_data.index, y='Net Value', label='Net Value')
    plt.title('Net Value Over Time')
    plt.xlabel('Date')
    plt.ylabel('Net Value (AED)')
    plt.legend()
    plt.savefig('static/net_value_plot.png')
    plt.close()

    return render_template('results.html', start_date=start_date, end_date=end_date)

# Run the app
if __name__ == '__main__':
    app.run(debug=True)
