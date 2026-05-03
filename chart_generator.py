import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from matplotlib.figure import Figure
import io
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
from database import DatabaseManager

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun', 'KaiTi']
plt.rcParams['axes.unicode_minus'] = False


class ChartGenerator:
    def __init__(self, db_manager: DatabaseManager = None):
        self.db_manager = db_manager or DatabaseManager()

    def create_operation_statistics_chart(self, days: int = 30) -> Figure:
        fig = Figure(figsize=(10, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        stats = self.db_manager.get_statistics()
        
        labels = ['总操作数', '排版操作', '总记录数', '成功数', '失败数']
        values = [
            stats['total_operations'],
            stats['total_format_operations'],
            stats['total_records'],
            stats['total_success'],
            stats['total_errors']
        ]
        
        colors = ['#3498db', '#2ecc71', '#9b59b6', '#1abc9c', '#e74c3c']
        
        bars = ax.bar(labels, values, color=colors, alpha=0.7)
        
        ax.set_title(f'操作统计概览 (最近{days}天)', fontsize=14, fontweight='bold')
        ax.set_ylabel('数量', fontsize=12)
        ax.set_xlabel('统计项目', fontsize=12)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2., height,
                    f'{int(height)}',
                    ha='center', va='bottom', fontsize=11)
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        fig.tight_layout()
        return fig

    def create_daily_operations_chart(self, days: int = 30) -> Figure:
        fig = Figure(figsize=(12, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        operations = self.db_manager.get_operations_by_date_range(
            start_date.strftime('%Y-%m-%d'),
            end_date.strftime('%Y-%m-%d 23:59:59')
        )
        
        daily_counts = {}
        for i in range(days + 1):
            date_key = (start_date + timedelta(days=i)).strftime('%Y-%m-%d')
            daily_counts[date_key] = {'generate': 0, 'format': 0, 'total': 0}
        
        for op in operations:
            created_at = op['created_at']
            if isinstance(created_at, str):
                try:
                    created_at = datetime.strptime(created_at, '%Y-%m-%d %H:%M:%S')
                except:
                    try:
                        created_at = datetime.strptime(created_at, '%Y-%m-%d')
                    except:
                        continue
            
            date_key = created_at.strftime('%Y-%m-%d')
            if date_key in daily_counts:
                daily_counts[date_key]['total'] += 1
                op_type = op.get('operation_type', '').lower()
                if '生成' in op_type or 'generate' in op_type:
                    daily_counts[date_key]['generate'] += 1
                elif '排版' in op_type or 'format' in op_type:
                    daily_counts[date_key]['format'] += 1
        
        dates = sorted(daily_counts.keys())
        generate_values = [daily_counts[d]['generate'] for d in dates]
        format_values = [daily_counts[d]['format'] for d in dates]
        
        x = range(len(dates))
        width = 0.35
        
        ax.bar([i - width/2 for i in x], generate_values, width, label='文档生成', color='#3498db', alpha=0.7)
        ax.bar([i + width/2 for i in x], format_values, width, label='排版操作', color='#2ecc71', alpha=0.7)
        
        ax.set_title(f'每日操作统计 (最近{days}天)', fontsize=14, fontweight='bold')
        ax.set_xlabel('日期', fontsize=12)
        ax.set_ylabel('操作次数', fontsize=12)
        ax.legend()
        
        if len(dates) > 10:
            step = len(dates) // 5
            ax.set_xticks(x[::step])
            ax.set_xticklabels([dates[i] for i in x[::step]], rotation=45)
        else:
            ax.set_xticks(x)
            ax.set_xticklabels(dates, rotation=45)
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        fig.tight_layout()
        return fig

    def create_operation_type_pie_chart(self) -> Figure:
        fig = Figure(figsize=(8, 8), dpi=100)
        ax = fig.add_subplot(111)
        
        stats = self.db_manager.get_statistics()
        operation_types = stats['operation_types']
        
        if not operation_types:
            ax.text(0.5, 0.5, '暂无数据', ha='center', va='center', fontsize=14)
            return fig
        
        labels = [item['operation_type'] for item in operation_types]
        sizes = [item['cnt'] for item in operation_types]
        
        colors = ['#3498db', '#2ecc71', '#9b59b6', '#e74c3c', '#f39c12', '#1abc9c']
        colors = colors[:len(labels)]
        
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
               startangle=90, shadow=True)
        ax.axis('equal')
        
        ax.set_title('操作类型分布', fontsize=14, fontweight='bold')
        
        return fig

    def create_success_rate_chart(self) -> Figure:
        fig = Figure(figsize=(10, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        stats = self.db_manager.get_statistics()
        
        total = stats['total_success'] + stats['total_errors']
        if total == 0:
            ax.text(0.5, 0.5, '暂无数据', ha='center', va='center', fontsize=14)
            return fig
        
        success_rate = stats['total_success'] / total * 100 if total > 0 else 0
        
        labels = ['成功', '失败']
        sizes = [stats['total_success'], stats['total_errors']]
        colors = ['#2ecc71', '#e74c3c']
        
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
               startangle=90, shadow=True)
        ax.axis('equal')
        
        ax.set_title(f'操作成功率 (成功率: {success_rate:.1f}%)', fontsize=14, fontweight='bold')
        
        return fig

    def create_weekly_operations_chart(self) -> Figure:
        fig = Figure(figsize=(10, 6), dpi=100)
        ax = fig.add_subplot(111)
        
        end_date = datetime.now()
        start_date = end_date - timedelta(days=60)
        
        operations = self.db_manager.get_operations_by_date_range(
            start_date.strftime('%Y-%m-%d'),
            end_date.strftime('%Y-%m-%d 23:59:59')
        )
        
        weekday_counts = {i: 0 for i in range(7)}
        weekday_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
        
        for op in operations:
            created_at = op['created_at']
            if isinstance(created_at, str):
                try:
                    created_at = datetime.strptime(created_at, '%Y-%m-%d %H:%M:%S')
                except:
                    continue
            
            weekday = created_at.weekday()
            weekday_counts[weekday] += 1
        
        x = range(7)
        values = [weekday_counts[i] for i in range(7)]
        
        bars = ax.bar(weekday_names, values, color='#3498db', alpha=0.7)
        
        ax.set_title('每周操作分布', fontsize=14, fontweight='bold')
        ax.set_xlabel('星期', fontsize=12)
        ax.set_ylabel('操作次数', fontsize=12)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2., height,
                    f'{int(height)}',
                    ha='center', va='bottom', fontsize=11)
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        fig.tight_layout()
        return fig

    def figure_to_bytes(self, fig: Figure) -> bytes:
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        return buf.getvalue()
