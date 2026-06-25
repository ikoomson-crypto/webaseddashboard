"""
Backend Logic - Data Management, Financial Analysis, Categories, PDF Report
"""

import pandas as pd
import json
import os
import io
from datetime import datetime
import xlsxwriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas


# ==================== PDF REPORT GENERATOR ====================
class PDFReportGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._create_custom_styles()

    def _create_custom_styles(self):
        """Create custom paragraph styles"""
        self.styles.add(ParagraphStyle(
            name='CustomReportTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1a237e'),
            alignment=TA_CENTER,
            spaceAfter=30
        ))

        self.styles.add(ParagraphStyle(
            name='ReportSectionHeader',
            parent=self.styles['Heading2'],
            fontSize=16,
            textColor=colors.HexColor('#0d47a1'),
            spaceAfter=12,
            spaceBefore=20
        ))

        self.styles.add(ParagraphStyle(
            name='ReportSubHeader',
            parent=self.styles['Heading3'],
            fontSize=14,
            textColor=colors.HexColor('#1565c0'),
            spaceAfter=8,
            spaceBefore=12
        ))

        self.styles.add(ParagraphStyle(
            name='ReportBodyText',
            parent=self.styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            alignment=TA_LEFT
        ))

        self.styles.add(ParagraphStyle(
            name='ReportMetricValue',
            parent=self.styles['Normal'],
            fontSize=12,
            textColor=colors.HexColor('#1b5e20'),
            alignment=TA_RIGHT,
            fontName='Helvetica-Bold'
        ))

        self.styles.add(ParagraphStyle(
            name='ReportMetricLabel',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.HexColor('#37474f'),
            alignment=TA_LEFT,
            fontName='Helvetica'
        ))

        self.styles.add(ParagraphStyle(
            name='ReportCommentText',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#424242'),
            alignment=TA_LEFT,
            leftIndent=20,
            spaceAfter=10
        ))

        self.styles.add(ParagraphStyle(
            name='ReportInsightTitle',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.HexColor('#bf360c'),
            alignment=TA_LEFT,
            fontName='Helvetica-Bold',
            spaceAfter=4
        ))

    def generate_report(self, df, report_data, filename):
        """Generate PDF report"""
        doc = SimpleDocTemplate(
            filename,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )

        story = []

        # TITLE
        story.append(Paragraph("Financial Performance Report", self.styles['CustomReportTitle']))
        story.append(Spacer(1, 0.1 * inch))
        story.append(Paragraph(
            f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}",
            self.styles['ReportBodyText']
        ))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1a237e')))
        story.append(Spacer(1, 0.2 * inch))

        # SECTION 1: SNAPSHOT
        story.append(Paragraph("1. SNAPSHOT", self.styles['ReportSectionHeader']))
        story.append(Paragraph(
            f"Period: {report_data.get('period', 'N/A')}",
            self.styles['ReportBodyText']
        ))
        story.append(Spacer(1, 0.1 * inch))

        snapshot_data = [
            ["Metric", "Amount"],
            ["Sales (Revenue)", f"${report_data['revenue']:,.2f}"],
            ["Cost of Sales (COS)", f"${report_data['cogs']:,.2f}"],
            ["Expenses", f"${report_data['expenses']:,.2f}"],
            ["Net Profit", f"${report_data['net_income']:,.2f}"]
        ]

        snapshot_table = Table(snapshot_data, colWidths=[3 * inch, 2 * inch])
        snapshot_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#bdbdbd')),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 11),
            ('TOPPADDING', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
            ('TEXTCOLOR', (1, 4), (1, 4),
             colors.HexColor('#1b5e20') if report_data['net_income'] > 0 else colors.HexColor('#c62828')),
            ('FONTNAME', (1, 4), (1, 4), 'Helvetica-Bold'),
        ]))
        story.append(snapshot_table)
        story.append(Spacer(1, 0.2 * inch))

        story.append(Paragraph(
            f"Gross Profit: ${report_data['gross_profit']:,.2f} | "
            f"Gross Margin: {report_data['gross_margin']:.1f}% | "
            f"Net Margin: {report_data['net_margin']:.1f}%",
            self.styles['ReportBodyText']
        ))
        story.append(Spacer(1, 0.2 * inch))

        # SECTION 2: PROFITABILITY ANALYSIS
        story.append(Paragraph("2. PROFITABILITY ANALYSIS", self.styles['ReportSectionHeader']))
        story.append(Spacer(1, 0.05 * inch))

        profit_data = [
            ["Metric", "Value", "Benchmark", "Status"],
            ["Gross Margin", f"{report_data['gross_margin']:.1f}%", "> 40%",
             "✅ Good" if report_data['gross_margin'] > 40 else "⚠️ Below Target"],
            ["Net Margin", f"{report_data['net_margin']:.1f}%", "> 15%",
             "✅ Good" if report_data['net_margin'] > 15 else "⚠️ Below Target"],
            ["Expense Ratio", f"{report_data['expense_ratio']:.1f}%", "< 30%",
             "✅ Good" if report_data['expense_ratio'] < 30 else "⚠️ High"],
        ]

        profit_table = Table(profit_data, colWidths=[1.5 * inch, 1.5 * inch, 1.5 * inch, 1.5 * inch])
        profit_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0d47a1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#bdbdbd')),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('ALIGN', (1, 1), (3, -1), 'CENTER'),
        ]))
        story.append(profit_table)
        story.append(Spacer(1, 0.15 * inch))

        story.append(Paragraph("Comments:", self.styles['ReportSubHeader']))
        for comment in report_data['profitability_comments']:
            story.append(Paragraph(f"• {comment}", self.styles['ReportCommentText']))
        story.append(Spacer(1, 0.2 * inch))

        # SECTION 3: LIQUIDITY ANALYSIS
        story.append(Paragraph("3. LIQUIDITY ANALYSIS", self.styles['ReportSectionHeader']))
        story.append(Spacer(1, 0.05 * inch))

        liquid_data = [
            ["Metric", "Value", "Benchmark", "Status"],
            ["Current Ratio", f"{report_data['current_ratio']:.2f}", "> 1.5",
             "✅ Good" if report_data['current_ratio'] > 1.5 else "⚠️ Low"],
            ["Quick Ratio", f"{report_data['quick_ratio']:.2f}", "> 1.0",
             "✅ Good" if report_data['quick_ratio'] > 1.0 else "⚠️ Low"],
            ["Working Capital", f"${report_data['working_capital']:,.0f}", "Positive",
             "✅ Good" if report_data['working_capital'] > 0 else "⚠️ Negative"],
        ]

        liquid_table = Table(liquid_data, colWidths=[1.5 * inch, 1.5 * inch, 1.5 * inch, 1.5 * inch])
        liquid_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0d47a1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#bdbdbd')),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('ALIGN', (1, 1), (3, -1), 'CENTER'),
        ]))
        story.append(liquid_table)
        story.append(Spacer(1, 0.15 * inch))

        story.append(Paragraph("Comments:", self.styles['ReportSubHeader']))
        for comment in report_data['liquidity_comments']:
            story.append(Paragraph(f"• {comment}", self.styles['ReportCommentText']))
        story.append(Spacer(1, 0.2 * inch))

        # SECTION 4: INSIGHTS AND SUGGESTIONS
        story.append(Paragraph("4. INSIGHTS AND SUGGESTIONS", self.styles['ReportSectionHeader']))
        story.append(Spacer(1, 0.05 * inch))

        for insight in report_data['insights']:
            story.append(Paragraph(f"📌 {insight['title']}", self.styles['ReportInsightTitle']))
            story.append(Paragraph(insight['message'], self.styles['ReportCommentText']))
            story.append(Paragraph(f"Action: {insight['action']}", self.styles['ReportBodyText']))
            story.append(Spacer(1, 0.05 * inch))

        # FOOTER
        story.append(Spacer(1, 0.3 * inch))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#e0e0e0')))
        story.append(Paragraph(
            "This report is auto-generated by the Financial Dashboard System.",
            self.styles['ReportBodyText']
        ))
        story.append(Paragraph(
            "All figures are based on the uploaded financial data.",
            self.styles['ReportBodyText']
        ))

        doc.build(story)
        return filename


# ==================== CATEGORY SERVICE ====================
class CategoryService:
    def __init__(self):
        self.categories = []
        self._load_categories()

    def _load_categories(self):
        """Load categories from file"""
        os.makedirs('data', exist_ok=True)
        if os.path.exists('data/categories.json'):
            try:
                with open('data/categories.json', 'r', encoding='utf-8') as f:
                    self.categories = json.load(f)
                    print(f"✅ Loaded {len(self.categories)} categories")
            except Exception as e:
                print(f"⚠️ Error loading categories: {e}")
                self.categories = []
        else:
            self.categories = []

    def _save_categories(self):
        """Save categories to file"""
        os.makedirs('data', exist_ok=True)
        with open('data/categories.json', 'w', encoding='utf-8') as f:
            json.dump(self.categories, f, indent=2, ensure_ascii=False)

    def get_categories(self):
        """Get all categories"""
        self._load_categories()
        return self.categories

    def add_category(self, name):
        """Add a new category"""
        self._load_categories()
        name = name.strip()
        if not name:
            return False, "Category name is required"
        if name in self.categories:
            return False, "Category already exists"

        self.categories.append(name)
        self._save_categories()
        return True, "Category added successfully"

    def delete_category(self, name):
        """Delete a category"""
        self._load_categories()
        if name not in self.categories:
            return False, "Category not found"

        self.categories.remove(name)
        self._save_categories()
        return True, "Category deleted successfully"


# ==================== DATA MANAGER ====================
class DataManager:
    def __init__(self):
        self.data = None
        self.transactions = []
        self._load_transactions_from_json()

    def _load_transactions_from_json(self):
        """Load transactions from JSON file and create dataframe"""
        os.makedirs('data', exist_ok=True)
        if os.path.exists('data/transactions.json'):
            try:
                with open('data/transactions.json', 'r', encoding='utf-8') as f:
                    self.transactions = json.load(f)
                    print(f"✅ Loaded {len(self.transactions)} transactions from JSON")

                    if self.transactions:
                        self.data = pd.DataFrame(self.transactions)
                        if 'date' in self.data.columns:
                            self.data['date'] = pd.to_datetime(self.data['date'])
                        print(f"✅ Created dataframe with {len(self.data)} rows")
                    return True
            except Exception as e:
                print(f"⚠️ Error loading transactions: {e}")
                return False
        else:
            print("ℹ️  No transactions.json file found")
            return False

    def load_from_file(self, file):
        """Load data from uploaded Excel or CSV file"""
        try:
            if file.filename.endswith('.xlsx'):
                df = pd.read_excel(file)
            elif file.filename.endswith('.csv'):
                df = pd.read_csv(file)
            else:
                return False, "Please upload Excel or CSV file"

            df = self._clean_data(df)
            self.data = df
            self._save_transactions_to_json()
            print(f"✅ Successfully loaded {len(df)} records")
            return True, f"Successfully loaded {len(df)} records"
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            return False, f"Error: {str(e)}"

    def _clean_data(self, df):
        """Clean and prepare data"""
        df.columns = df.columns.str.strip()

        column_mapping = {
            'Account Name': 'account_name',
            'Category': 'category',
            'Date': 'date',
            'Entity': 'entity',
            'Amount': 'amount'
        }

        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df.rename(columns={old_name: new_name}, inplace=True)

        if 'date' in df.columns:
            df['date'] = pd.to_datetime(df['date'], errors='coerce')

        if 'amount' in df.columns:
            df['amount'] = pd.to_numeric(df['amount'], errors='coerce')

        required_columns = ['amount', 'date', 'category', 'entity']
        for col in required_columns:
            if col in df.columns:
                df = df.dropna(subset=[col])

        return df

    def _save_transactions_to_json(self):
        """Save transactions to JSON file"""
        os.makedirs('data', exist_ok=True)
        if self.data is not None and not self.data.empty:
            transactions = self.data.to_dict('records')
            for t in transactions:
                if 'date' in t and pd.notna(t['date']):
                    t['date'] = t['date'].strftime('%Y-%m-%d')
                if 'amount' in t and pd.notna(t['amount']):
                    t['amount'] = float(t['amount'])

            with open('data/transactions.json', 'w', encoding='utf-8') as f:
                json.dump(transactions, f, indent=2, ensure_ascii=False)

            self.transactions = transactions
            print(f"✅ Saved {len(transactions)} transactions to JSON")

    def get_data(self):
        return self.data

    def get_transactions(self):
        return self.transactions

    def has_data(self):
        return self.data is not None and not self.data.empty

    def delete_transactions(self, indices):
        """Delete transactions by index"""
        if self.data is None or self.data.empty:
            return False

        valid_indices = [i for i in indices if i < len(self.data)]
        if not valid_indices:
            return False

        try:
            self.data = self.data.drop(index=valid_indices).reset_index(drop=True)
            self._save_transactions_to_json()
            return True
        except Exception as e:
            print(f"Error deleting transactions: {e}")
            return False

    def clear_data(self):
        """Clear all data"""
        self.data = None
        self.transactions = []
        if os.path.exists('data/transactions.json'):
            os.remove('data/transactions.json')
            print("✅ Cleared transactions.json")
        return True

    def export_to_excel(self, df=None):
        """Export data to Excel with multiple sheets"""
        if df is None:
            df = self.data

        if df is None or df.empty:
            return None

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df = df.copy()
            export_df.rename(columns={
                'account_name': 'Account Name',
                'category': 'Category',
                'date': 'Date',
                'entity': 'Entity',
                'amount': 'Amount'
            }, inplace=True)

            export_df.to_excel(writer, sheet_name='Raw Data', index=False)

            analyzer = FinancialAnalyzer()
            analyzer.data = df
            kpis = analyzer.calculate_kpis(df)

            summary_data = {
                'Metric': ['Total Revenue', 'Total Expenses', 'Net Income', 'Profit Margin', 'Total Records'],
                'Value': [
                    f"${kpis['revenue']:,.2f}",
                    f"${kpis['expenses']:,.2f}",
                    f"${kpis['net_income']:,.2f}",
                    f"{kpis['profit_margin']:.2f}%",
                    len(df)
                ]
            }

            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        output.seek(0)
        return output

    def export_template(self):
        """Export empty Excel template"""
        output = io.BytesIO()
        category_service = CategoryService()
        categories = category_service.get_categories()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            template_data = {
                'Account Name': ['Enter account name here', '', '', '', ''],
                'Category': [categories[0] if categories else 'Enter category', '', '', '', ''],
                'Date': [datetime.now().strftime('%Y-%m-%d'), '', '', '', ''],
                'Entity': ['Enter entity name', '', '', '', ''],
                'Amount': [0, 0, 0, 0, 0]
            }

            template_df = pd.DataFrame(template_data)
            template_df.to_excel(writer, sheet_name='Data Template', index=False)

            instructions = {
                'Column': ['Account Name', 'Category', 'Date', 'Entity', 'Amount'],
                'Description': [
                    'Name of the specific account',
                    'Category (e.g., Sales, Cost of Sales, Expenses)',
                    'Transaction date in YYYY-MM-DD format',
                    'Entity/Department/Region name',
                    'Transaction amount'
                ],
                'Example': ['Product Sales', 'Sales', '2024-01-15', 'North', '50000']
            }

            instructions_df = pd.DataFrame(instructions)
            instructions_df.to_excel(writer, sheet_name='Instructions', index=False)

            worksheet = writer.sheets['Data Template']
            header_format = writer.book.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white'})

            for col_num, value in enumerate(template_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:B', 25)
            worksheet.set_column('C:C', 15)
            worksheet.set_column('D:D', 18)
            worksheet.set_column('E:E', 15)

        output.seek(0)
        return output


# ==================== CATEGORY TYPE MANAGER ====================
class CategoryTypeManager:
    def __init__(self):
        self.category_types_file = 'data/category_types.json'
        self.revenue_categories = ['Sales']
        self.expense_categories = ['Cost of Sales', 'Expenses']
        self._load_category_types()

    def _load_category_types(self):
        """Load category type mappings from file"""
        os.makedirs('data', exist_ok=True)
        if os.path.exists(self.category_types_file):
            try:
                with open(self.category_types_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.revenue_categories = data.get('revenue_categories', ['Sales'])
                    self.expense_categories = data.get('expense_categories', ['Cost of Sales', 'Expenses'])
                    print(
                        f"✅ Loaded category types: Revenue={self.revenue_categories}, Expense={self.expense_categories}")
            except Exception as e:
                print(f"⚠️ Error loading category types: {e}")
        else:
            self._save_category_types()

    def _save_category_types(self):
        """Save category type mappings to file"""
        os.makedirs('data', exist_ok=True)
        try:
            with open(self.category_types_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'revenue_categories': self.revenue_categories,
                    'expense_categories': self.expense_categories
                }, f, indent=2, ensure_ascii=False)
            print(f"✅ Saved category types")
        except Exception as e:
            print(f"⚠️ Error saving category types: {e}")

    def get_revenue_categories(self):
        return self.revenue_categories

    def get_expense_categories(self):
        return self.expense_categories

    def update_revenue_categories(self, categories):
        self.revenue_categories = categories
        self._save_category_types()

    def update_expense_categories(self, categories):
        self.expense_categories = categories
        self._save_category_types()


# ==================== FINANCIAL ANALYZER ====================
class FinancialAnalyzer:
    def __init__(self):
        self.data = None
        # CATEGORY-ONLY COMPUTATION - Exact match only
        self.revenue_categories = ['Sales']
        self.expense_categories = ['Cost of Sales', 'Expenses']
        self.ignored_categories = ['Non-Current Assets', 'Current Assets', 'Current Liabilities', 'Owners Equity']

    def calculate_kpis(self, df):
        """
        Calculate KPIs using ONLY CATEGORIES - EXACT MATCH
        Revenue = Categories that EXACTLY MATCH revenue_categories
        Expenses = Categories that EXACTLY MATCH expense_categories
        """
        if df is None or df.empty:
            return {
                'revenue': 0,
                'expenses': 0,
                'net_income': 0,
                'profit_margin': 0,
                'has_data': False
            }

        # EXACT MATCH for revenue - No partial matching
        revenue_mask = df['category'].isin(self.revenue_categories)
        revenue = df[revenue_mask]['amount'].sum()

        # EXACT MATCH for expenses - No partial matching
        expense_mask = df['category'].isin(self.expense_categories)
        expenses = abs(df[expense_mask]['amount'].sum())

        net_income = revenue - expenses
        profit_margin = (net_income / revenue * 100) if revenue > 0 else 0

        return {
            'revenue': float(revenue),
            'expenses': float(expenses),
            'net_income': float(net_income),
            'profit_margin': float(profit_margin),
            'has_data': True
        }

    def get_entity_performance(self, df):
        """Get performance by entity using EXACT CATEGORY MATCH"""
        if df is None or df.empty:
            return []

        entity_data = []
        for entity in df['entity'].unique():
            entity_df = df[df['entity'] == entity]

            # EXACT MATCH for revenue
            revenue = entity_df[entity_df['category'].isin(self.revenue_categories)]['amount'].sum()

            # EXACT MATCH for expenses
            expense_mask = entity_df['category'].isin(self.expense_categories)
            expenses = abs(entity_df[expense_mask]['amount'].sum())

            net_profit = revenue - expenses
            profit_margin = (net_profit / revenue * 100) if revenue > 0 else 0

            entity_data.append({
                'entity': str(entity),
                'revenue': float(revenue),
                'expenses': float(expenses),
                'net_profit': float(net_profit),
                'profit_margin': float(profit_margin)
            })

        return entity_data

    def get_category_distribution(self, df):
        """Get distribution by category"""
        if df is None or df.empty:
            return {}
        distribution = df.groupby('category')['amount'].sum().abs().to_dict()
        return {str(k): float(v) for k, v in distribution.items()}

    def get_trend_analysis(self, df):
        """Get trend analysis using EXACT CATEGORY MATCH"""
        if df is None or df.empty:
            return {'revenue_trend': [], 'expense_trend': [], 'has_data': False}

        df_copy = df.copy()
        df_copy['month'] = df_copy['date'].dt.to_period('M')
        df_copy['year_month'] = df_copy['month'].astype(str)

        revenue_trend = []
        expense_trend = []

        for month in df_copy['year_month'].unique():
            month_data = df_copy[df_copy['year_month'] == month]

            # EXACT MATCH for revenue
            revenue = month_data[month_data['category'].isin(self.revenue_categories)]['amount'].sum()

            # EXACT MATCH for expenses
            expense_mask = month_data['category'].isin(self.expense_categories)
            expenses = abs(month_data[expense_mask]['amount'].sum())

            revenue_trend.append({'date': str(month), 'amount': revenue})
            expense_trend.append({'date': str(month), 'amount': expenses})

        return {
            'revenue_trend': revenue_trend,
            'expense_trend': expense_trend,
            'has_data': True
        }

    def generate_insights(self, df):
        """Generate actionable insights"""
        if df is None or df.empty:
            return [{
                'type': 'info',
                'title': 'No Data Loaded',
                'message': 'Please upload your financial data to see insights',
                'action': 'Go to Transactions page to upload data'
            }]

        insights = []
        kpis = self.calculate_kpis(df)

        if kpis['profit_margin'] < 10:
            insights.append({
                'type': 'critical',
                'title': 'Low Profit Margin',
                'message': f'Net profit margin is {kpis["profit_margin"]:.1f}% - below healthy benchmark',
                'action': 'Review cost structure and optimize pricing strategy'
            })
        elif kpis['profit_margin'] < 20:
            insights.append({
                'type': 'warning',
                'title': 'Moderate Profitability',
                'message': f'Profit margin at {kpis["profit_margin"]:.1f}% - room for improvement',
                'action': 'Focus on operational efficiency and cost reduction'
            })
        else:
            insights.append({
                'type': 'success',
                'title': 'Strong Profitability',
                'message': f'Excellent profit margin of {kpis["profit_margin"]:.1f}%',
                'action': 'Maintain current strategy and look for expansion opportunities'
            })

        entity_perf = self.get_entity_performance(df)
        if entity_perf:
            best_entity = max(entity_perf, key=lambda x: x['net_profit'])
            insights.append({
                'type': 'info',
                'title': 'Top Performer',
                'message': f'{best_entity["entity"]} leads with ${best_entity["net_profit"]:,.0f} profit',
                'action': 'Analyze and replicate ' + best_entity['entity'] + "'s successful strategies"
            })

        return insights

    def get_report_data(self, df):
        """Generate comprehensive report data using the selected end date for closing balances"""
        if df is None or df.empty:
            return None

        kpis = self.calculate_kpis(df)

        # ===== PROFIT & LOSS (Income Statement) - Summed for the period =====
        revenue_mask = df['category'].isin(self.revenue_categories)
        revenue = df[revenue_mask]['amount'].sum()

        cogs_mask = df['category'].isin(['Cost of Sales'])
        cogs = abs(df[cogs_mask]['amount'].sum())

        expense_mask = df['category'].isin(self.expense_categories)
        expenses = abs(df[expense_mask]['amount'].sum())

        gross_profit = revenue - cogs
        gross_margin = (gross_profit / revenue * 100) if revenue > 0 else 0
        net_margin = (kpis['net_income'] / revenue * 100) if revenue > 0 else 0
        expense_ratio = (expenses / revenue * 100) if revenue > 0 else 0

        # ===== BALANCE SHEET (Closing Balances) - Get balance at END DATE =====
        # Get the end date from the filtered data (this is the selected end date)
        end_date = df['date'].max()

        def get_closing_balance_at_date(category_name, end_date):
            """Get the closing balance for a balance sheet category at the end date"""
            # Get all transactions for this category up to the end date
            category_df = df[df['category'] == category_name]
            if category_df.empty:
                return 0

            # Filter transactions up to the end date
            category_df = category_df[category_df['date'] <= end_date]
            if category_df.empty:
                return 0

            # Sort by date and get the last transaction (closing balance)
            category_df = category_df.sort_values('date')
            latest = category_df.iloc[-1]
            return abs(latest['amount'])

        # Get closing balances for balance sheet items at the end date
        current_assets = get_closing_balance_at_date('Current Assets', end_date)
        non_current_assets = get_closing_balance_at_date('Non-Current Assets', end_date)
        current_liabilities = get_closing_balance_at_date('Current Liabilities', end_date)
        owners_equity = get_closing_balance_at_date('Owners Equity', end_date)

        total_assets = current_assets + non_current_assets
        current_ratio = current_assets / current_liabilities if current_liabilities > 0 else 0
        quick_ratio = current_ratio * 0.8
        working_capital = current_assets - current_liabilities

        period = f"{df['date'].min().strftime('%Y-%m-%d')} to {df['date'].max().strftime('%Y-%m-%d')}"

        profitability_comments = self._get_profitability_comments(kpis, gross_margin, net_margin, expense_ratio)
        liquidity_comments = self._get_liquidity_comments(current_ratio, quick_ratio, working_capital)
        insights = self.generate_insights(df)

        return {
            'period': period,
            'revenue': kpis['revenue'],
            'cogs': cogs,
            'expenses': kpis['expenses'],
            'net_income': kpis['net_income'],
            'gross_profit': gross_profit,
            'gross_margin': gross_margin,
            'net_margin': net_margin,
            'expense_ratio': expense_ratio,
            'current_assets': current_assets,
            'non_current_assets': non_current_assets,
            'total_assets': total_assets,
            'current_liabilities': current_liabilities,
            'owners_equity': owners_equity,
            'current_ratio': current_ratio,
            'quick_ratio': quick_ratio,
            'working_capital': working_capital,
            'profitability_comments': profitability_comments,
            'liquidity_comments': liquidity_comments,
            'insights': insights
        }

    def _get_profitability_comments(self, kpis, gross_margin, net_margin, expense_ratio):
        """Generate profitability comments"""
        comments = []

        if gross_margin > 40:
            comments.append(
                "Gross margin is strong at " + f"{gross_margin:.1f}%, indicating good product profitability and efficient cost management.")
        elif gross_margin > 30:
            comments.append(
                "Gross margin at " + f"{gross_margin:.1f}% is moderate. Consider optimizing production costs or adjusting pricing.")
        else:
            comments.append(
                "Gross margin at " + f"{gross_margin:.1f}% is below benchmark. Review cost of sales and pricing strategy.")

        if net_margin > 15:
            comments.append(
                "Net profit margin of " + f"{net_margin:.1f}% is excellent. Operating efficiency is well maintained.")
        elif net_margin > 10:
            comments.append(
                "Net profit margin of " + f"{net_margin:.1f}% is good but has room for improvement. Focus on operational efficiency.")
        else:
            comments.append(
                "Net profit margin of " + f"{net_margin:.1f}% is concerning. Immediate cost reduction measures are recommended.")

        if expense_ratio < 30:
            comments.append(
                "Expense ratio at " + f"{expense_ratio:.1f}% is well controlled. Cost structure is efficient.")
        elif expense_ratio < 40:
            comments.append(
                "Expense ratio at " + f"{expense_ratio:.1f}% is moderate. Review operational costs for potential savings.")
        else:
            comments.append(
                "Expense ratio at " + f"{expense_ratio:.1f}% is high. Detailed cost analysis and reduction initiatives needed.")

        return comments

    def _get_liquidity_comments(self, current_ratio, quick_ratio, working_capital):
        """Generate liquidity comments"""
        comments = []

        if current_ratio > 2.0:
            comments.append(
                "Current ratio of " + f"{current_ratio:.2f} indicates excellent short-term liquidity position.")
        elif current_ratio > 1.5:
            comments.append(
                "Current ratio of " + f"{current_ratio:.2f} indicates adequate liquidity to meet short-term obligations.")
        else:
            comments.append(
                "Current ratio of " + f"{current_ratio:.2f} is below benchmark. Monitor working capital management.")

        if quick_ratio > 1.0:
            comments.append(
                "Quick ratio of " + f"{quick_ratio:.2f} shows strong ability to meet immediate obligations without selling inventory.")
        else:
            comments.append(
                "Quick ratio of " + f"{quick_ratio:.2f} suggests reliance on inventory. Consider improving collection of receivables.")

        if working_capital > 0:
            comments.append(
                "Positive working capital of $" + f"{working_capital:,.0f} indicates healthy operational liquidity.")
        else:
            comments.append(
                "Negative working capital of $" + f"{working_capital:,.0f}. Immediate attention needed to address liquidity issues.")

        return comments


# ==================== EXPORT FUNCTIONS ====================
# These are kept for backward compatibility

def export_to_excel(df):
    """Export data to Excel - calls DataManager.export_to_excel"""
    dm = DataManager()
    dm.data = df
    return dm.export_to_excel(df)


def export_template():
    """Export template - calls DataManager.export_template"""
    dm = DataManager()
    return dm.export_template()