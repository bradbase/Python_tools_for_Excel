
import logging
from datetime import datetime

from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator

logging.basicConfig(level=logging.INFO)

beginning = datetime.now()
print("loading file")
filename = r'Nested_sum.xlsx'
compiler = ModelCompiler()
print("model compiler made", datetime.now() - beginning)
new_model = compiler.read_and_parse_archive(filename)
# print("new_model.cells", new_model.cells)
print("read_and_parse_archive took", datetime.now() - beginning)
new_model.build_code()
print("build_code took", datetime.now() - beginning)

print("now evaluating")
evaluator = Evaluator(new_model)
print("evaluator made", datetime.now() - beginning)

columns = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
# columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U']
# columns = ['V', 'W', 'X', 'Y', 'Z']
# columns = ['H']

max_rows = 3
# max_rows = 3
addresses = ['Sheet1!{}{}'.format(column, row) for row in range(2, max_rows) for column in columns]

print("addresses made", datetime.now() - beginning)
second_beginning = datetime.now()
for address in addresses:
    evaluated_value = evaluator.evaluate(address)
    # new_model.set_cell_value(address, evaluated_value)
    # print("eval_ref.cache_info()", evaluator.eval_ref.cache_info())
    print("EVALUATED VALUE", address, evaluated_value)

# print("used cached_values", evaluator.cache_count, "of", len(evaluator.evaluated_cells), "evaluated cells at a ratio cache_hits:evaluated_cells", evaluator.cache_count/len(evaluator.evaluated_cells) if len(evaluator.evaluated_cells) != 0 else 0)
# print("evaluate.cache_info", evaluator.evaluate.cache_info())
# print("eval_ref.cache_info()", evaluator.eval_ref.cache_info(), evaluator.cache_count)
print("Evaluation done", datetime.now() - second_beginning)
print("all done", datetime.now() - beginning)
print()
print()

# third_beginning = datetime.now()
# for address in addresses:
#     evaluated_value = evaluator.evaluate(address)
#     new_model.set_cell_value(address, evaluated_value)
#     # print("eval_ref.cache_info()", evaluator.eval_ref.cache_info())
#     print("EVALUATED VALUE", address, evaluated_value)
# print("third evaluation done", datetime.now() - third_beginning)
