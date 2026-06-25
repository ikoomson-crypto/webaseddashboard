"""
Financial Dashboard - Main Application
"""

from flask import Flask, render_template, request, jsonify, session, send_file, Response
from datetime import datetime
import os
import pandas as pd
import tempfile
import json

from backend import DataManager, FinancialAnalyzer, CategoryService, PDFReportGenerator, CategoryTypeManager

# Initialize Flask app
app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-in-production'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# Create upload folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('data', exist_ok=True)

# Initialize services
data_manager = DataManager()
analyzer = FinancialAnalyzer()
category_service = CategoryService()
pdf_generator = PDFReportGenerator()
category_type_manager = CategoryTypeManager()

# Update analyzer with saved category types
analyzer.revenue_categories = category_type_manager.get_revenue_categories()
analyzer.expense_categories = category_type_manager.get_expense_categories()


# ==================== ROUTES ====================

@app.route('/')
def index():
    """Dashboard - Analytics only"""
    data = data_manager.get_data()
    has_data = data_manager.has_data()
    return render_template('dashboard.html', has_data=has_data, data=data if data is not None else [])


@app.route('/transactions')
def transactions_page():
    """Transactions - Upload, view, delete, export"""
    transactions = data_manager.get_transactions()
    return render_template('transactions.html', transactions=transactions)


@app.route('/reports')
def reports_page():
    """Reports - Export and view reports"""
    data = data_manager.get_data()
    has_data = data_manager.has_data()
    return render_template('reports.html', has_data=has_data, data=data if data is not None else [])


@app.route('/categories')
def categories_page():
    """Categories - Manage categories"""
    categories = category_service.get_categories()
    return render_template('categories.html', categories=categories)


@app.route('/pdf-report')
def pdf_report_page():
    """PDF Report - View and export to PDF"""
    data = data_manager.get_data()
    has_data = data_manager.has_data()

    if not has_data:
        return render_template('pdf_report.html', has_data=False, report_data=None)

    # Get filters from URL parameters
    entity = request.args.get('entity', 'All')
    start_date = request.args.get('start')
    end_date = request.args.get('end')

    # Filter data
    filtered = data.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    # Get report data
    report_data = analyzer.get_report_data(filtered)

    return render_template('pdf_report.html', has_data=True, report_data=report_data)


# ==================== EXPORT PDF ROUTE ====================

@app.route('/api/export-pdf')
def export_pdf():
    """Generate and download PDF report"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({'error': 'No data loaded'}), 400

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    entity = request.args.get('entity', 'All')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    if filtered.empty:
        return jsonify({'error': 'No data available for selected filters'}), 400

    report_data = analyzer.get_report_data(filtered)

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    pdf_generator.generate_report(filtered, report_data, temp_file.name)
    temp_file.close()

    filename = f"financial_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

    return send_file(
        temp_file.name,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/generate-report')
def generate_report():
    """Generate report data for selected period"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({'success': False, 'message': 'No data loaded'}), 400

    entity = request.args.get('entity', 'All')
    start_date = request.args.get('start')
    end_date = request.args.get('end')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    if filtered.empty:
        return jsonify({'success': False, 'message': 'No data available for selected filters'}), 400

    report_data = analyzer.get_report_data(filtered)

    return jsonify({'success': True, 'report_data': report_data})


# ==================== CATEGORY TYPE MANAGEMENT ROUTES ====================

@app.route('/api/category-types', methods=['GET'])
def get_category_types():
    """Get the current category type mappings"""
    return jsonify({
        'revenue_categories': analyzer.revenue_categories,
        'expense_categories': analyzer.expense_categories
    })


@app.route('/api/category-types/update', methods=['POST'])
def update_category_types():
    """Update which categories are Revenue and Expense"""
    data = request.json
    revenue = data.get('revenue_categories', [])
    expense = data.get('expense_categories', [])

    # Update the analyzer
    analyzer.revenue_categories = revenue
    analyzer.expense_categories = expense

    # Save to persistent storage
    category_type_manager.update_revenue_categories(revenue)
    category_type_manager.update_expense_categories(expense)

    return jsonify({
        'success': True,
        'message': 'Category types updated successfully',
        'revenue_categories': analyzer.revenue_categories,
        'expense_categories': analyzer.expense_categories
    })


# ==================== API ROUTES ====================

@app.route('/upload', methods=['POST'])
def upload_file():
    """Upload data file"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'No file selected'})

    ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
    allowed = {'xlsx', 'xls', 'csv'}
    if ext not in allowed:
        return jsonify({'success': False, 'message': f'Invalid file type. Allowed: {", ".join(allowed)}'})

    success, message = data_manager.load_from_file(file)
    if success:
        session['data_loaded'] = True
        return jsonify({'success': True, 'message': message})
    else:
        return jsonify({'success': False, 'message': message})


@app.route('/clear-data')
def clear_data():
    """Clear all data"""
    data_manager.clear_data()
    session.pop('data_loaded', None)
    return jsonify({'success': True, 'message': 'Data cleared successfully'})


@app.route('/api/kpis')
def get_kpis():
    """Get KPI metrics"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({'error': 'No data loaded', 'has_data': False})

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    entity = request.args.get('entity', 'All')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    kpis = analyzer.calculate_kpis(filtered)
    return jsonify(kpis)


@app.route('/api/trends')
def get_trends():
    """Get trend data"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({'error': 'No data loaded', 'has_data': False})

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    entity = request.args.get('entity', 'All')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    if filtered.empty:
        return jsonify({'revenue_trend': [], 'expense_trend': [], 'has_data': False})

    trends = analyzer.get_trend_analysis(filtered)
    return jsonify(trends)


@app.route('/api/category-distribution')
def get_category_distribution():
    """Get category distribution"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({})

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    entity = request.args.get('entity', 'All')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    distribution = analyzer.get_category_distribution(filtered)
    return jsonify(distribution)


@app.route('/api/entity-performance')
def get_entity_performance():
    """Get entity performance"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify([])

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]

    performance = analyzer.get_entity_performance(filtered)
    return jsonify(performance)


@app.route('/api/insights')
def get_insights():
    """Get insights"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify([])

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]

    insights = analyzer.generate_insights(filtered)
    return jsonify(insights)


@app.route('/api/entities')
def get_entities():
    """Get list of entities"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify([])
    return jsonify(df['entity'].unique().tolist())


@app.route('/api/export-data')
def export_data():
    """Export data as CSV"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({'error': 'No data loaded'})

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    entity = request.args.get('entity', 'All')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    csv_data = filtered.to_csv(index=False)

    response = Response(csv_data, mimetype='text/csv')
    response.headers.set("Content-Disposition", "attachment",
                         filename=f"financial_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    return response


@app.route('/api/export-excel')
def export_excel():
    """Export data as Excel"""
    df = data_manager.get_data()
    if not data_manager.has_data():
        return jsonify({'error': 'No data loaded'}), 400

    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    entity = request.args.get('entity', 'All')

    filtered = df.copy()
    if start_date:
        filtered = filtered[filtered['date'] >= pd.to_datetime(start_date)]
    if end_date:
        filtered = filtered[filtered['date'] <= pd.to_datetime(end_date)]
    if entity != 'All' and 'entity' in filtered.columns:
        filtered = filtered[filtered['entity'] == entity]

    if filtered.empty:
        return jsonify({'error': 'No data available for selected filters'}), 400

    excel_file = data_manager.export_to_excel(filtered)
    if excel_file is None:
        return jsonify({'error': 'Failed to export data'}), 400

    filename = f"financial_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/download-template')
def download_template():
    """Download Excel template"""
    excel_file = data_manager.export_template()
    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='financial_data_template.xlsx'
    )


@app.route('/api/transactions/delete', methods=['POST'])
def delete_transactions():
    """Delete selected transactions"""
    try:
        data = request.json
        indices = data.get('indices', [])

        if not indices:
            return jsonify({'success': False, 'message': 'No transactions selected'})

        success = data_manager.delete_transactions(indices)
        if success:
            return jsonify({'success': True, 'message': f'Deleted {len(indices)} transaction(s)'})
        else:
            return jsonify({'success': False, 'message': 'Failed to delete transactions'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/categories/add', methods=['POST'])
def add_category():
    """Add a new category"""
    data = request.json
    name = data.get('category_name', '')
    success, message = category_service.add_category(name)
    return jsonify({'success': success, 'message': message})


@app.route('/api/categories/delete/<name>', methods=['DELETE'])
def delete_category(name):
    """Delete a category"""
    success, message = category_service.delete_category(name)
    return jsonify({'success': success, 'message': message})


if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("   FINANCIAL DASHBOARD SYSTEM")
    print("=" * 60)
    print("\n📁 Files:")
    print("   - app.py (routes)")
    print("   - backend.py (logic)")
    print("\n📊 CATEGORY-ONLY COMPUTATION ENABLED")
    print("   - Revenue = Categories matching: " + ', '.join(analyzer.revenue_categories))
    print("   - Expenses = Categories matching: " + ', '.join(analyzer.expense_categories))
    print("   - Amount signs are IGNORED")
    print("\n🌐 Starting web server...")
    print("\n👉 Open your browser and go to: http://localhost:5000")
    print("=" * 60)
    print("\nAvailable pages:")
    print("   📊 Dashboard: http://localhost:5000")
    print("   📋 Transactions: http://localhost:5000/transactions")
    print("   📄 PDF Report: http://localhost:5000/pdf-report")
    print("   📈 Reports: http://localhost:5000/reports")
    print("   🏷️  Categories: http://localhost:5000/categories")
    print("\n" + "=" * 60 + "\n")

    app.run(debug=True, host='0.0.0.0', port=5000)