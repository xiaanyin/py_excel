# py_excel
Run python in excel vba using xlwings.

--python環境整備
python -m venv env
env\Scripts\activate.bat

--インストール
pip install xlwings

--アドイン
xlwings addin install

--必須ではない依存関係
--これらのパッケージは必須ではありませんが、xlwingsと強力に連携することからインストールを強く推奨します
pip install "xlwings[all]"
pip install --upgrade xlwings

--アンインストール
xlwings addin remove
pip uninstall xlwings
