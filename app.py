import csv
import io
from datetime import datetime
from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# フォントの登録
pdfmetrics.registerFont(TTFont('NotoSansJP', 'fonts/NotoSansJP-Regular.ttf'))

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    # 初期値として None を設定
    profit_or_loss = None
    asset = {
        '現金': 0.0,
        '売掛金': 0.0,
        '在庫': 0.0,
        '不動産': 0.0,
        '設備': 0.0,
    }
    liability = {
        '短期借入金': 0.0,
        '買掛金': 0.0,
        '長期借入金': 0.0,
    }
    equity = 0.0
    # 損益計算書の項目初期化
    sales = 0.0
    other_revenue = 0.0
    cost_of_sales = 0.0
    selling_expenses = 0.0
    admin_expenses = 0.0
    interest_expense = 0.0
    tax_expense = 0.0

    if request.method == 'POST':
        try:
            # フォームデータの取得
            sales = float(request.form.get('sales', 0.0))
            other_revenue = float(request.form.get('other_revenue', 0.0))
            cost_of_sales = float(request.form.get('cost_of_sales', 0.0))
            selling_expenses = float(request.form.get('selling_expenses', 0.0))
            admin_expenses = float(request.form.get('admin_expenses', 0.0))
            interest_expense = float(request.form.get('interest_expense', 0.0))
            tax_expense = float(request.form.get('tax_expense', 0.0))

            asset['現金'] = float(request.form.get('cash', 0.0))
            asset['売掛金'] = float(request.form.get('accounts_receivable', 0.0))
            asset['在庫'] = float(request.form.get('inventory', 0.0))
            asset['不動産'] = float(request.form.get('property', 0.0))
            asset['設備'] = float(request.form.get('equipment', 0.0))
            liability['短期借入金'] = float(request.form.get('short_term_loan', 0.0))
            liability['買掛金'] = float(request.form.get('accounts_payable', 0.0))
            liability['長期借入金'] = float(request.form.get('long_term_loan', 0.0))
            equity = float(request.form.get('equity', 0.0))

            # 損益計算書
            profit_or_loss = (sales + other_revenue) - (cost_of_sales + selling_expenses + admin_expenses + interest_expense + tax_expense)

            # 貸借対照表の計算
            total_assets = sum(asset.values())  # 総資産
            total_liabilities = sum(liability.values())  # 総負債
            total_liabilities_and_equity = total_liabilities + equity  # 負債・資本の合計

            # actionをチェック
            # CSVまたはPDFの出力ボタンが押された場合
            action = request.form.get('action')
            if action == 'csv':
                return generate_csv(sales, other_revenue, cost_of_sales, selling_expenses, admin_expenses, interest_expense, tax_expense, asset, liability, equity)
            elif action == 'pdf':
                return generate_pdf(sales, other_revenue, cost_of_sales, selling_expenses, admin_expenses, interest_expense, tax_expense, asset, liability, equity)

            return render_template(
                'index.html',
                sales=sales,
                other_revenue=other_revenue,
                cost_of_sales=cost_of_sales,
                selling_expenses=selling_expenses,
                admin_expenses=admin_expenses,
                interest_expense=interest_expense,
                tax_expense=tax_expense,
                profit_or_loss=profit_or_loss,
                asset=asset,
                liability=liability,
                equity=equity,
                total_assets=total_assets,
                total_liabilities=total_liabilities,
                total_liabilities_and_equity=total_liabilities_and_equity
            )
        except Exception as e:
            print(f"Error: {e}")
            # エラーが発生した場合、エラーメッセージを表示する
            error_message = "データの送信に失敗しました。もう一度お試しください。"
            return render_template('index.html', error_message=error_message)
    # GETリクエスト時には空のフォームを表示
    return render_template('index.html', profit_or_loss=profit_or_loss)  # 初期値として profit_or_loss を渡す

# 会計帳簿を直接編集するページ
@app.route('/accounting')
def accounting():
    return render_template('accounting.html')

# 管理ページ
@app.route('/admin')
def admin():
    return render_template('admin.html')

# csv export
def generate_csv(sales, other_revenue, cost_of_sales, selling_expenses, admin_expenses, interest_expense, tax_expense, asset, liability, equity):
    output = io.StringIO()
    writer = csv.writer(output)

    # 損益計算書
    writer.writerow(["損益計算書"])
    writer.writerow(["売上高", int(sales)])
    writer.writerow(["その他収益", int(other_revenue)])
    writer.writerow(["売上原価", int(cost_of_sales)])
    writer.writerow(["販売費", int(selling_expenses)])
    writer.writerow(["一般管理費", int(admin_expenses)])
    writer.writerow(["支払利息", int(interest_expense)])
    writer.writerow(["税金", int(tax_expense)])
    writer.writerow(["損益", int(sales + other_revenue - (cost_of_sales + selling_expenses + admin_expenses + interest_expense + tax_expense))])

    # 貸借対照表
    writer.writerow(["貸借対照表"])
    writer.writerow(["現金", int(asset['現金'])])
    writer.writerow(["売掛金", int(asset['売掛金'])])
    writer.writerow(["在庫", int(asset['在庫'])])
    writer.writerow(["不動産", int(asset['不動産'])])
    writer.writerow(["設備", int(asset['設備'])])
    writer.writerow(["総資産", int(sum(asset.values()))])

    writer.writerow(["短期借入金", int(liability['短期借入金'])])
    writer.writerow(["買掛金", int(liability['買掛金'])])
    writer.writerow(["長期借入金", int(liability['長期借入金'])])
    writer.writerow(["総負債", int(sum(liability.values()))])

    writer.writerow(["資本", int(equity)])
    writer.writerow(["負債・資本の合計", int(sum(liability.values()) + equity)])

    # エンコードして送信
    csv_data = output.getvalue()
    bom_utf8 = '\ufeff' + csv_data 
    byte_io = io.BytesIO()
    byte_io.write(bom_utf8.encode('utf-8'))
    byte_io.seek(0)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"accounting_data_{timestamp}.csv"

    return send_file(byte_io, as_attachment=True, download_name=filename, mimetype="text/csv")

# PDF出力処理
def generate_pdf(sales, other_revenue, cost_of_sales, selling_expenses, admin_expenses, interest_expense, tax_expense, asset, liability, equity):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont("NotoSansJP", 12)

    c.drawString(100, 750, f"損益計算書")
    c.drawString(100, 730, f"売上高: ¥{int(sales)}")
    c.drawString(100, 710, f"その他収益: ¥{int(other_revenue)}")
    c.drawString(100, 690, f"売上原価: ¥{int(cost_of_sales)}")
    c.drawString(100, 670, f"販売費: ¥{int(selling_expenses)}")
    c.drawString(100, 650, f"一般管理費: ¥{int(admin_expenses)}")
    c.drawString(100, 630, f"支払利息: ¥{int(interest_expense)}")
    c.drawString(100, 610, f"税金: ¥{int(tax_expense)}")
    c.drawString(100, 590, f"損益: ¥{int(sales + other_revenue - (cost_of_sales + selling_expenses + admin_expenses + interest_expense + tax_expense))}")

    # 空行を追加
    c.drawString(100, 570, "")
    c.drawString(100, 550, "貸借対照表")

    c.drawString(100, 530, f"現金: ¥{int(asset['現金'])}")
    c.drawString(100, 510, f"売掛金: ¥{int(asset['売掛金'])}")
    c.drawString(100, 490, f"在庫: ¥{int(asset['在庫'])}")
    c.drawString(100, 470, f"不動産: ¥{int(asset['不動産'])}")
    c.drawString(100, 450, f"設備: ¥{int(asset['設備'])}")
    c.drawString(100, 430, f"総資産: ¥{int(sum(asset.values()))}")

    c.drawString(100, 410, f"短期借入金: ¥{int(liability['短期借入金'])}")
    c.drawString(100, 390, f"買掛金: ¥{int(liability['買掛金'])}")
    c.drawString(100, 370, f"長期借入金: ¥{int(liability['長期借入金'])}")
    c.drawString(100, 350, f"総負債: ¥{int(sum(liability.values()))}")

    c.drawString(100, 330, f"資本: ¥{int(equity)}")
    c.drawString(100, 310, f"負債・資本の合計: ¥{int(sum(liability.values()) + equity)}")

    c.showPage()
    c.save()

    buffer.seek(0)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"accounting_data_{timestamp}.pdf"

    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="application/pdf")


if __name__ == '__main__':
    app.run(debug=True)
