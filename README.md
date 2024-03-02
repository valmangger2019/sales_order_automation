# sales_order_automation
A code to create sales orders in the SAP system automatically, based on information from an excel file.


#Importing all libraries
import pandas as pd
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import win32com.client
import sys
import openpyxl


