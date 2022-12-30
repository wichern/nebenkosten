init:
    pip install -r requirements.txt

test:
    python test_createbill.py

.PHONY: init test
