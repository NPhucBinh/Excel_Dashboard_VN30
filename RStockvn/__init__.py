# Copyright 2023 Nguyen Phuc Binh @ GitHub
# See LICENSE for details.
__version__ = "2.1.10.1"
__author__ ="Nguyen Phuc Binh"
__copyright__ = "Copyright 2023, Nguyen Phuc Binh"
__license__ = "MIT"
__email__ = "nguyenphucbinh67@gmail.com"
__website__ = "https://github.com/NPhucBinh"



from .update_package import (check_for_updates,updates_package_RStockvn)
from . import user_agent
from .chrome_driver.chromedriver_setup import (check_var, remove_file_old)
from .stockvn import (get_foreign_historical_vnd,get_price_historical_vnd,key_id,list_company,update_company,report_finance_vnd,report_finance_cf,
    getCPI_vietstock,solieu_sanxuat_congnghiep,solieu_banle_vietstock,solieu_XNK_vietstock,solieu_FDI_vietstock,tygia_vietstock,solieu_tindung_vietstock,laisuat_vietstock,
    solieu_danso_vietstock,solieu_GDP_vietstock,get_data_result_order,get_info_cp,momentum_ck)

from selenium import webdriver


remove_file_old()
check_var()