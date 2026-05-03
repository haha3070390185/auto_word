import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Any, Optional
from config import DATABASE_PATH


class DatabaseManager:
    def __init__(self, db_path: str = DATABASE_PATH):
        self.db_path = db_path
        self.init_database()

    def get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_database(self):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS operations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                operation_type TEXT NOT NULL,
                template_name TEXT,
                output_path TEXT,
                record_count INTEGER,
                success_count INTEGER,
                error_count INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'completed',
                details TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS generated_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                operation_id INTEGER,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                file_size INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (operation_id) REFERENCES operations (id)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS format_operations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                operation_type TEXT NOT NULL,
                files_count INTEGER,
                font_name TEXT,
                font_size REAL,
                line_spacing REAL,
                indentation REAL,
                bold TEXT,
                italic TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'completed',
                details TEXT
            )
        ''')
        
        conn.commit()
        conn.close()

    def add_operation(self, operation_type: str, template_name: str = None,
                       output_path: str = None, record_count: int = 0,
                       success_count: int = 0, error_count: int = 0,
                       status: str = 'completed', details: str = None) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO operations (operation_type, template_name, output_path,
                                    record_count, success_count, error_count, status, details)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (operation_type, template_name, output_path, record_count,
              success_count, error_count, status, details))
        
        operation_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return operation_id

    def add_generated_file(self, operation_id: int, file_name: str,
                           file_path: str, file_size: int = 0):
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO generated_files (operation_id, file_name, file_path, file_size)
            VALUES (?, ?, ?, ?)
        ''', (operation_id, file_name, file_path, file_size))
        
        conn.commit()
        conn.close()

    def add_format_operation(self, operation_type: str, files_count: int,
                             font_name: str = None, font_size: float = None,
                             line_spacing: float = None, indentation: float = None,
                             bold: str = None, italic: str = None,
                             status: str = 'completed', details: str = None) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO format_operations (operation_type, files_count, font_name,
                                           font_size, line_spacing, indentation, bold,
                                           italic, status, details)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (operation_type, files_count, font_name, font_size, line_spacing,
              indentation, bold, italic, status, details))
        
        operation_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return operation_id

    def get_operations(self, limit: int = 100) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM operations ORDER BY created_at DESC LIMIT ?
        ''', (limit,))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

    def get_format_operations(self, limit: int = 100) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM format_operations ORDER BY created_at DESC LIMIT ?
        ''', (limit,))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

    def get_operations_by_type(self, operation_type: str, limit: int = 100) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM operations WHERE operation_type = ? ORDER BY created_at DESC LIMIT ?
        ''', (operation_type, limit))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

    def get_generated_files(self, operation_id: int) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM generated_files WHERE operation_id = ?
        ''', (operation_id,))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

    def get_operations_by_date_range(self, start_date: str, end_date: str) -> List[Dict]:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM operations 
            WHERE created_at >= ? AND created_at <= ? 
            ORDER BY created_at DESC
        ''', (start_date, end_date))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

    def get_statistics(self) -> Dict[str, Any]:
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT COUNT(*) FROM operations')
        total_operations = cursor.fetchone()[0]
        
        cursor.execute('SELECT COUNT(*) FROM format_operations')
        total_format_operations = cursor.fetchone()[0]
        
        cursor.execute('SELECT SUM(record_count) FROM operations')
        total_records = cursor.fetchone()[0] or 0
        
        cursor.execute('SELECT SUM(success_count) FROM operations')
        total_success = cursor.fetchone()[0] or 0
        
        cursor.execute('SELECT SUM(error_count) FROM operations')
        total_errors = cursor.fetchone()[0] or 0
        
        cursor.execute('''
            SELECT operation_type, COUNT(*) as cnt 
            FROM operations 
            GROUP BY operation_type 
            ORDER BY cnt DESC
        ''')
        operation_types = [dict(row) for row in cursor.fetchall()]
        
        conn.close()
        
        return {
            'total_operations': total_operations,
            'total_format_operations': total_format_operations,
            'total_records': total_records,
            'total_success': total_success,
            'total_errors': total_errors,
            'operation_types': operation_types
        }
