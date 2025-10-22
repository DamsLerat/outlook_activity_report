# outlook_activity_report

Languages:\
![Python](https://img.shields.io/badge/Python%203-FFD43B?style=for-the-badge&logo=python&logoColor=blue "Python 3.11")

OS:\
![Debian](https://img.shields.io/badge/Debian-A81D33?style=for-the-badge&logo=debian&logoColor=white "Debian 12")
![Linux](https://img.shields.io/badge/Linux-FCC624?style=for-the-badge&logo=linux&logoColor=black "GNU/Linux")

License:\
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Author:\
![Author](https://badgen.net/badge/icon/Damien%20Lerat?icon=buymeacoffee&label)

## 💡 Introduction

Create activity report from mails/meeting/jira/commits

## 🗒️ Getting Started

### 🛠️ Installation

Clone the repository

  ```shell
  git clone https://github.com/DamsLerat/outlook_activity_report
  cd outlook_activity_report
  ```

### 📦 Dependencies

This project use:

- Python 3.11 and venv
- git for cloning and updating other repositories

### 📦 Install required packages

Install Required Debian packages: \

  ```shell
    sudo apt install  git python3-venv python3.11
  ```

__Create__  or __Activate__ the virtual environment (In the repository root folder).\
Note: setup_venv.sh also install required  packages dependencies using packages provided in *requirements.txt* \

  ```shell
    cd outlook_activity_report
    ./setup_venv.sh

  ```

### 💻 Usage

  ```shell
  source DSA-venv//bin/activate
  ./OAR-venv/bin/python ./extract_activity.py
  ```
