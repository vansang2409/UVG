"""Chạy đối chiếu nhóm doanh thu từ dữ liệu của dự án."""
from pathlib import Path
import argparse, os, subprocess, sys
ROOT=Path(__file__).resolve().parent
parser=argparse.ArgumentParser(description='Đối chiếu nhóm AMIS/meInvoice; không phát hành hóa đơn.')
parser.add_argument('--data-dir',type=Path,default=ROOT/'data')
parser.add_argument('--output-dir',type=Path,default=ROOT/'outputs')
parser.add_argument('--validate',action='store_true')
args=parser.parse_args()
required={
 'Du lieu Amis':['So_chi_tiet_ban_hang 2024.xlsx','So_chi_tiet_ban_hang T1-6.2025.xlsx'],
 'Du lieu Meinvoice':['Bảng kê chi tiết hóa đơn 2024.xls','Bảng kê chi tiết hóa đơn 2025.xls','Bang_ke_hoa_don_da_su_dung_2024-T6.2025.xls','Bang_ke_hoa_don_da_su_dung_T6.2025.xls','Bang_ke_hoa_don_da_su_dung_T6.2025 L2.xls','DS bị điều chỉnh 2024.xls','DS bị điều chỉnh 2025.xls','DS bị thay thế 2024.xls'],
}
missing=[args.data_dir/folder/name for folder,names in required.items() for name in names if not (args.data_dir/folder/name).is_file()]
if missing:
 print('Thiếu file đầu vào:',file=sys.stderr)
 for p in missing: print(p,file=sys.stderr)
 sys.exit(2)
if args.validate:
 print('Đủ 10 file đầu vào. Chưa xác nhận dữ liệu đầy đủ so với CQT.')
 sys.exit(0)
try:
 import openpyxl
except ImportError:
 print('Cài thư viện: python -m pip install -r requirements.txt',file=sys.stderr)
 sys.exit(2)
env=os.environ.copy()
env['UVG_DATA_DIR']=str(args.data_dir.resolve())
env['UVG_OUTPUT_DIR']=str(args.output_dir.resolve())
env['PYTHONUTF8']='1'
sys.exit(subprocess.call([sys.executable,str(ROOT/'scripts'/'reconcile_uvg_groups.py')],cwd=ROOT,env=env))
