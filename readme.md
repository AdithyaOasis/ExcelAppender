Steps to run -
Download Python
Make sure python is present in the environment path
Go to cmd, and cd into the EXCELAPPENDER folder and type the following commands sequentially -
venv env
env\scripts\activate
pip install -r requirements.txt

Now the setup is ready!

Keep the current Main sheet in Current_Main folder as "Main.xlsx".
Keep the to-be added bill in Bills folder as "new_bill.xlsx".

Now back in the terminal, run the commands -
env\scripts\activate
py app.py

The updated sheet will be available in the upd folder!
