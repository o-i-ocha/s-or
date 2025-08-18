from flask import Flask, render_template, request

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

if __name__ == '__main__':
    app.run(debug=True)
