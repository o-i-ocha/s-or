import csv
import io
from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

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
    return render_template('index.html', profit_or_loss=profit_or_loss)

# CSV出力処理
def download_csv(sales, other_revenue, cost_of_sales, selling_expenses, admin_expenses, interest_expense, tax_expense, asset, liability, equity):
    profit_or_loss = (sales + other_revenue) - (cost_of_sales + selling_expenses + admin_expenses + interest_expense + tax_expense)
    total_assets = sum(asset.values())
    total_liabilities = sum(liability.values())
    total_liabilities_and_equity = total_liabilities + equity

    # CSV出力
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["損益計算書"])
    writer.writerow(["売上高", sales])
    writer.writerow(["その他収益", other_revenue])
    writer.writerow(["売上原価", cost_of_sales])
    writer.writerow(["販売費", selling_expenses])
    writer.writerow(["一般管理費", admin_expenses])
    writer.writerow(["支払利息", interest_expense])
    writer.writerow(["税金", tax_expense])
    writer.writerow(["損益", profit_or_loss])

    writer.writerow([])
    writer.writerow(["貸借対照表"])
    writer.writerow(["現金", asset['現金']])
    writer.writerow(["売掛金", asset['売掛金']])
    writer.writerow(["在庫", asset['在庫']])
    writer.writerow(["不動産", asset['不動産']])
    writer.writerow(["設備", asset['設備']])
    writer.writerow(["総資産", total_assets])

    writer.writerow(["短期借入金", liability['短期借入金']])
    writer.writerow(["買掛金", liability['買掛金']])
    writer.writerow(["長期借入金", liability['長期借入金']])
    writer.writerow(["総負債", total_liabilities])

    writer.writerow(["資本", equity])
    writer.writerow(["負債・資本の合計", total_liabilities_and_equity])

    output.seek(0)
    return send_file(output, as_attachment=True, download_name="financial_statement.csv", mimetype="text/csv")


# PDF出力処理
def download_pdf(sales, other_revenue, cost_of_sales, selling_expenses, admin_expenses, interest_expense, tax_expense, asset, liability, equity):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.drawString(100, 750, f"損益計算書")
    c.drawString(100, 730, f"売上高: ¥{sales}")
    c.drawString(100, 710, f"その他収益: ¥{other_revenue}")
    c.drawString(100, 690, f"売上原価: ¥{cost_of_sales}")
    c.drawString(100, 670, f"販売費: ¥{selling_expenses}")
    c.drawString(100, 650, f"一般管理費: ¥{admin_expenses}")
    c.drawString(100, 630, f"支払利息: ¥{interest_expense}")
    c.drawString(100, 610, f"税金: ¥{tax_expense}")
    c.drawString(100, 590, f"損益: ¥{sales + other_revenue - (cost_of_sales + selling_expenses + admin_expenses + interest_expense + tax_expense)}")
    
    # 空行を追加
    c.drawString(100, 570, "")
    c.drawString(100, 550, "貸借対照表")

    # 資産の項目
    c.drawString(100, 530, f"現金: ¥{asset['現金']}")
    c.drawString(100, 510, f"売掛金: ¥{asset['売掛金']}")
    c.drawString(100, 490, f"在庫: ¥{asset['在庫']}")
    c.drawString(100, 470, f"不動産: ¥{asset['不動産']}")
    c.drawString(100, 450, f"設備: ¥{asset['設備']}")
    c.drawString(100, 430, f"総資産: ¥{sum(asset.values())}")

    # 負債の項目
    c.drawString(100, 410, f"短期借入金: ¥{liability['短期借入金']}")
    c.drawString(100, 390, f"買掛金: ¥{liability['買掛金']}")
    c.drawString(100, 370, f"長期借入金: ¥{liability['長期借入金']}")
    c.drawString(100, 350, f"総負債: ¥{sum(liability.values())}")

    # 資本の項目
    c.drawString(100, 330, f"資本: ¥{equity}")
    c.drawString(100, 310, f"負債・資本の合計: ¥{sum(liability.values()) + equity}")

    c.showPage()
    c.save()

    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="financial_statement.pdf", mimetype="application/pdf")


if __name__ == '__main__':
    app.run(debug=True)
