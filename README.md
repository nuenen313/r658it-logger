# r658it-logger

## Prerequisites

- Python 3.14  
- Git  
- Poetry  
- GPIB interface with drivers (Keysight IO Libraries or NI-VISA)

---

## Install Python 3.14

1. Download Python 3.14 from:  
   https://www.python.org/downloads/

2. Run the installer  
   - Enable "Add Python to PATH"  
   - Complete installation  

3. Verify:

    python --version

---

## Install Poetry

Install using pip:

    pip install poetry

Or using the official installer:

    (Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | python -

Verify:

    poetry --version

---

## Clone Repository

    git clone https://github.com/nuenen313/r658it-logger.git
    cd r658it-logger

---

## Install Dependencies

    poetry install

---

## Use

Activate environment:

    poetry shell

Run:

    python main.py

---

## Notes

- Ensure GPIB drivers are installed and working  
- Confirm the instrument is visible in the connection tool  
- Check that the GPIB address matches the configuration  
