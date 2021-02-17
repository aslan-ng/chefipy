from unit_convertor.objects import UnitManager
import os


BASE_DIR = os.getcwd()

UNIT_DIR = os.path.join(BASE_DIR, "unit_convertor")
um = UnitManager(UNIT_DIR)
print(um.convert(2, 'liter', 'cup'))