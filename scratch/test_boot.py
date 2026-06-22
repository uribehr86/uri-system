import sys
import io
import time

sys.stdout.reconfigure(encoding='utf-8', write_through=True)

print("1. Importing flask...")
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
print("2. Importing psycopg2...")
import psycopg2
from psycopg2 import pool
from psycopg2.extras import RealDictCursor
print("3. Importing sqlite3...")
import sqlite3
print("4. Importing genai...")
try:
    from google import genai
    print("google.genai imported")
except ImportError:
    try:
        import google.generativeai as genai
        print("google.generativeai imported")
    except ImportError:
        genai = None
        print("genai is None")

print("5. Initializing GenAI Client...")
import os
from dotenv import load_dotenv
load_dotenv()
try:
    if genai and hasattr(genai, 'Client'):
        print("Calling genai.Client...")
        genai_client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))
        print("genai.Client created")
    else:
        print("No genai.Client class found")
except Exception as e:
    print(f"GenAI init failed: {e}")

print("6. Creating Flask App...")
app = Flask(__name__)
print("7. Success! Boot completed without blocking.")
