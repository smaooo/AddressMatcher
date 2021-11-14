# This is a prgram to compare the address list of guilds and occupations distributed in a geographical area
# and append the object id of the matched address from the address list of that geographical area to the guilds list.
# The address list of each geographical area contains the name of the streets and passages in that area with their
# geographical cooridinates. By using this program each guild or occupation will obtain a geographical coordinate 
# to use in the GIS (Geographic Information System).
# This program is specifically designed for identification and matching of the persian names and data.
# The first final veision and phase of this program is tested with the distribution of guilds and occupations in Tehran.

from openpyxl import load_workbook
from multiprocessing import Process, Manager
import time, os, datetime
from hazm import Normalizer, word_tokenize
from persian import convert_ar_characters as arconvert


class Sheet:
    # A sheet respresents a specific sheet of a xlsx workbook with sensitivity to rows, columns and cells of that workbook

    def __init__(self, name):
        # Using openxlpy import the xlsx (ms office excel) workbook and set its primary sheet as active for analysis
        self._wb = load_workbook(name + '.xlsx')
        self._sheet_wb = self._wb.active
        
    
    def value_calc(self, col, row):
        # Return the value of the cell in the intersection of the given column and row
        self._value = self._sheet_wb[col + str(row)].value
        return self._value

    def row_counter(self):
        # Calculate the last row number which has any value except None
        return len(self._sheet_wb['A'])+1
    
    def save_wb(self, name):
        # Save the workbook with the given name
        self._wb.save(name + '.xlsx')

    def value_merger(self, given_sheet_list):
        # Merge the values in the given list with the current workbook
        for sub_list in given_sheet_list:
            self._sheet_wb['K' + str(sub_list[0])].value = sub_list[1]
            self._sheet_wb['L' + str(sub_list[0])].value = sub_list[2]
            self._sheet_wb['M' + str(sub_list[0])].value = sub_list[3]

class Timer:
    # Create a Timer which can return the remaining time of a process
    # and the elapsed time

    def __init__(self, total):
        # Get the time at the initializing moment and define the total value of the process
        self._start = datetime.datetime.now()
        self._total = total

    def remains(self, done):
        # Calculate the remaining time until the end of the process by calculating the duration
        # of the finished sub-tasks and predicting the remaining sub-tasks duration.

        now = datetime.datetime.now() # Get the current moment time to compare with the begining moment.
        left = (self._total - done) * (now - self._start) / done # Calculate the reamining time
        sec = int(left.total_seconds())
        # If the remaining time is less than 1 minute, return the time in seconds
        if sec < 60:    
            return '{} seconds'.format(sec)
        # If the remaining time is bigger than a minute and less than one hour, return the time in minutes
        elif 60 <= sec < 3600:
            return '{} minutes'.format(int(sec / 60))
        # If the remaining time is bigger than one hour, return the time in hours
        else:
            return '{} hours'.format(int(sec / 3600))
    
    def now_time(self):
        return datetime.datetime.now()

    def calc_remaining(self, begin_time, current_row, total):
        now = datetime.datetime.now()
        left =  (now - begin_time) * (total - current_row)
        sec = int(left.total_seconds())
        if sec < 60:    
            return '{} seconds'.format(sec)
        # If the remaining time is bigger than a minute and less than one hour, return the time in minutes
        elif 60 <= sec < 3600:
            return '{} minutes'.format(int(sec / 60))
        # If the remaining time is bigger than one hour, return the time in hours
        else:
            return '{} hours'.format(int(sec / 3600))
    

    def elapsed(self, desired_time):
        # Calculate the elapsed time since the begining of the process

        now = datetime.datetime.now() # Get the current moment time
        elapsed = now - self._start # calculate the elapsed time
        # Check if the elapsed time equals to the desired time
        if desired_time == 0:
            return elapsed
        elif int(elapsed.total_seconds()) % desired_time == 0: 
            return True
        else:
            return False

class Address:
    # An Address in a normalized address with removed extra spaces 
    # and arabic characters converted to persian characters

    def __init__(self):
        self._normalizer = Normalizer()
    def persian_corrector(self, lit_input):
        if lit_input != None:
            return arconvert(self._normalizer.normalize(lit_input))
        
class Matching_data:
    # This is a class to store the most repeating variables such as lists 
    # to reduce the amount of repetition and complexity

    def __init__(self, input_sheet, control_sheet, col_value_list):
        self._input_sheet = input_sheet # store the guilds sheet
        self._control_sheet = control_sheet # store the address sheet
        self._col_value_list = col_value_list # store the matched list

    def get_input_sheet(self):
        # Return the guild list 
        return self._input_sheet
    
    def get_control_sheet(self):
        # Return the address list
        return self._control_sheet
    
    def get_col_value_list(self):
        # Return the mathced list
        return self._col_value_list
        
def main():
    input_sheet = Sheet(r'D:\Nima2\Mostaghelat') # Define the workbook representing the list of the guilds and occupations
    control_sheet = Sheet(r'D:\Nima2\All_Points') # Define the workbook representing the list of the passages of a geographical area
    t_checker(input_sheet, control_sheet) 

def list_looker(the_list, index_input):
    # Check if the specified row with the index_input number has matched with an address and coordination
    
    # If the list is not empty check if the row number (index_input) is in the matched list or not
    if len(the_list) > 0:
        for sub_list in the_list:
            # If the row number isn't the the matched list, return True to compare that row with the general address list
            if sub_list[0] != index_input:
                return True
            
    # If the list is still empty, return True to compare the lists with each other
    else:
        return True


def matcher(matching_data, address_list, index_input, context):
    # Detemine if the specific address at the row with the index_control number is matched with one of the
    # addresses provided for the guilds and occupaotions.
    # This determination is based on the type of the passage. The function should find the specific type of the
    # passage described by the context and after that look for a matching address in the occupation list.
    
    address_corrector = Address()
    for index_control in range(2, 15991):

        if matching_data.get_input_sheet().value_calc('J', index_input) and matching_data.get_control_sheet().value_calc('E', index_control) != None:
            # Correct the text value of the both cells of the occupation list and the passage list.
            control_corrected = address_corrector.persian_corrector(matching_data.get_control_sheet().value_calc('E', index_control)) + ' '
            # Check if the value of the address list is same as the specified type by the context.
            if matching_data.get_control_sheet().value_calc('D', index_control) == context[0] or matching_data.get_control_sheet().value_calc('D', index_control) == context[1]:
                # Check if the row at the index input is previously matched with another address. If so pass this row.
                if list_looker (matching_data.get_col_value_list(), index_input) == True:
                    # Check if the selected address is present in the occupation address.
                    if control_corrected in address_list[1]:
                        # If both addresses are matched add the obeject id of the matched address to the selected row of the occupation list
                        matching_data.get_col_value_list().append([index_input, matching_data.get_control_sheet().value_calc('A', index_control), 'T', context[0]])
                        return True

    


def word_matcher(matching_data, address_list, index_input, context):
    address_corrector = Address()
    for index_control in range(2, 15991):
        if matching_data.get_input_sheet().value_calc('J', index_input) and matching_data.get_control_sheet().value_calc('E', index_control) != None:
            control_corrected = address_corrector.persian_corrector(matching_data.get_control_sheet().value_calc('E', index_control))
            if matching_data.get_control_sheet().value_calc('D', index_control) == context[0] or matching_data.get_control_sheet().value_calc('D', index_control) == context[1]:
                if list_looker (matching_data.get_col_value_list(), index_input) == True:
                    if len(address_list[2]) > 1 and control_corrected == (address_list[2][0] + ' ' + address_list[2][1]):
                        matching_data.get_col_value_list().append([index_input, matching_data.get_control_sheet().value_calc('A', index_control), 'T', context[0]])
                        return True

                    elif len(address_list[2]) > 2 and control_corrected == (address_list[2][0] + ' ' + address_list[2][1] + ' ' + address_list[2][2]):
                        matching_data.get_col_value_list().append([index_input, matching_data.get_control_sheet().value_calc('A', index_control), 'T', context[0]])
                        return True

                    elif control_corrected == address_list[0]:
                        matching_data.get_col_value_list().append([index_input, matching_data.get_control_sheet().value_calc('A', index_control), 'T', context[0]])
                        return True
                    
                    
def manager(matching_data, input_range_start, input_range_end):
    # Manager should control the both lists and select rows from them and send the selected row numbers to the matcher function.

    # Contexts are determining the priority of the passage types that should be matched with occupations' addresses
    timer = Timer(input_range_end) 

    for index_input in range(input_range_start, input_range_end):
        #contexts = [['primary', 'primary_link', True], ['secondary', 'secondary_link', False], ['tertiary', 'teritiary_link', False],
        #['service', 'service', False], ['residential', 'residential', False]]
        begin_time = timer.now_time()
        keep_looking = True
        address_corrector = Address()
        input_corrected = address_corrector.persian_corrector(matching_data.get_input_sheet().value_calc('J', index_input))
        address_list = [word_tokenize(input_corrected)[0], input_corrected, word_tokenize(input_corrected)]
        for address in address_list:
            context_index = 0
            contexts = [['primary', 'primary_link', True], ['secondary', 'secondary_link', False], ['tertiary', 'teritiary_link', False],
            ['service', 'service', False], ['residential', 'residential', False]]
        # If there isn't a matching option for that occupation, keep looking until an address is matched or there isn't 
        # a mtching option
            while keep_looking:
                if address_list.index(address) == 0 and keep_looking:
                    
                    # Check if it is allowed to look for the specific passage type
                    if contexts[context_index][2] == True and keep_looking:    
                        word_matched = word_matcher(matching_data, address_list, index_input, contexts[context_index])
                        if word_matched == True:
                            keep_looking = False
                            break
                        elif context_index < len(contexts) -1:
                            context_index +=1
                            contexts[context_index][2] = True
                        elif word_matched != True and context_index == len(contexts) - 1:                              
                            break                               
                    else:
                        keep_looking = False
                        break
                elif address_list.index(address) == 1 and keep_looking:
                    # Check if it is allowed to look for the specific passage type
                    if contexts[context_index][2] == True and keep_looking:  
                         # Use matcher function to determine if the address is matching with the occupation address
                        matched = matcher(matching_data, address_list, index_input, contexts[context_index])
                                # If address is matched with the occupation break the loop and get the next row for evaluation
                        if matched == True:
                            keep_looking = False
                            break
                            # If it is the last passage type (service), and there isn't any matching address add to the attempt number
                            # and repete the process

                                # If the addresses with thhe current passage type isn't matching with the current occupation, switch to the next passage type
                        elif context_index < len(contexts) -1:
                            context_index +=1
                            contexts[context_index][2] = True
                        elif matched != True and context_index == len(contexts) - 1:     
                            keep_looking = False                         
                            break 
                        else:
                            keep_looking = False
                            break
        print(timer.calc_remaining(begin_time, index_input, input_range_end))

        prev_save = 0
        if len(matching_data.get_col_value_list()) % 100 == 0 and len(matching_data.get_col_value_list()) != prev_save:
            prev_save = len(matching_data.get_col_value_list())
            autosave(matching_data)
def autosave(matching_data):
    # Merge the data from result list with the occupation list
    matching_data.get_input_sheet().value_merger(matching_data.get_col_value_list())
    # Save the merged occupation list
    matching_data.get_input_sheet().save_wb('newtest')
    print('Auto Save')

def time_based_save(matching_data, timer):
    if timer.elapsed(10*60):
        autosave(matching_data)

def t_checker(input_sheet, control_sheet):
    # This function should initialize the prerequisites for the matching process.
    # The matching process runes in a concurrency and several CPU cores will be assigned
    # to divide the workload and enhance the performance.

    if __name__ == '__main__':
        # Create a timer
        timer = Timer(input_sheet.row_counter()) 

        high_index = input_sheet.row_counter() # Get the highest row number of the occupation list
        low_index = 2 # Set the lowest row number as 2 (first row is the header and should not be evaluated)
        
        # Create a range list to divide the rows of the occupation list between different CPU cores
        index_range_list = [low_index, high_index] 
        local_range = 0
        # Create a global list to store all of the matched data from concurrent process
        col_value_list = Manager().list()
        #col_value_list = []
        # Use matching_data class to store the repetitive variables
        matching_data = Matching_data(input_sheet, control_sheet, col_value_list)

        # Based on the number of cpu cores set the ranges for the number of occupation list rows
        for cores in range(os.cpu_count()):    
            local_range = local_range + high_index/(os.cpu_count()-2)
            index_range_list.insert(cores+1, int(local_range))
        
        # Create concurrent processes for matching function
        pool = [Process(target=manager, args=(matching_data, index_range_list[p], index_range_list[p+1])) for p in range(os.cpu_count()-2)]
        #auto_save_pool = Process(target=time_based_save, args=(matching_data, timer))
        # Tell the process worker pools to start
        #auto_save_pool.start()
        for p in pool:
            p.start()

        # Tell the process worker pools to join            
        #auto_save_pool.join()
        for p in pool:
            p.join()
        time.sleep(0.2)
        #manager(matching_data, 2, 10)
        # Merge the final result list with the occupations list
        input_sheet.value_merger(col_value_list)
        # Save the occupations list
        input_sheet.save_wb('newtest')

main()
