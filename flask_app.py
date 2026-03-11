import os
import sqlite3
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify
from dotenv import load_dotenv
import openai

# טעינת הגדרות
load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'uri_system_2026_final')
DB_PATH = 'system_data.db'

# הגדרת ה-AI
openai.api_key = os.getenv("OPENAI_API_KEY")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def index():
    # דף הנחיתה המעוצב שראינו בתמונה
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = request.form.get('username', '').strip()
        pwd = request.form.get('password', '').strip()
        
        # כניסה מהירה לאורי
        if user == 'admin_uri' and pwd == 'uri*':
            session.update({'logged_in': True, 'username': user})
            return redirect(url_for('dashboard'))
            
        flash('שם משתמש או סיסמה שגויים', 'danger')
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'): return redirect(url_for('login'))
    
    conn = get_db_connection()
    # משיכת הנתונים שסנכרנת הרגע (ה-2056)
    total = conn.execute('SELECT COUNT(*) FROM computers').fetchone()[0]
    recent = conn.execute('SELECT * FROM computers ORDER BY rowid DESC LIMIT 10').fetchall()
    conn.close()
    
    return render_template('dashboard.html', total=total, recent=recent)

@app.route('/ask_ai', methods=['POST'])
def ask_ai():
    if not session.get('logged_in'): return jsonify({'answer': 'נא להתחבר.'})
    
    user_query = request.json.get('query', '')
    conn = get_db_connection()
    count = conn.execute('SELECT COUNT(*) FROM computers').fetchone()[0]
    conn.close()
    
    try:
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": f"אתה העוזר של אורי. במערכת יש {count} מחשבים. תענה בעברית קצרה."},
                {"role": "user", "content": user_query}
            ]
        )
        answer = response.choices[0].message.content
    except:
        answer = f"אורי, ה-AI כרגע לא זמין, אבל בסיס הנתונים מחובר ויש בו {count} מחשבים."
        
    return jsonify({'answer': answer})

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    print("\n" + "="*30)
    print("🚀 URI SYSTEM IS LIVE!")
    print("🌐 Link: http://127.0.0.1:5000")
    print("="*30 + "\n")
    app.run(debug=True, port=5000)