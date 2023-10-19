import logging
from dataclasses import dataclass


logger = logging.getLogger('config')



template = '''
{
    "sheets_name" : "PRODUCT_CODES",
    "sheets_name_done" : "DONE_LIST",
    "sheets_id" : "1SkeiwNjCXAOHo1qQaIUn-66FwC4lQg2TWXaNfRAYQLw",
    "sheets_range" : "A2:K",
    "sheets_range_done" : "A2:C",
    "try_number" : 3,
    "values_list" : [
        "code",
        "description",
        "type",
        "price_1",
        "price_2",
        "price_3",
        "price_4",
        "price_5",
        "price_6",
        "repos_cost"
    ]
}

'''


@dataclass
class Config_Values:
    sheets_name : str
    sheets_name_done : str
    sheets_id : str
    sheets_range : str
    sheets_range_done : str
    try_number : int
    values_list : list

    def __eq__(self, __o: object) -> bool:
        if not isinstance(__o, self.__class__):
            return NotImplemented
        else:
            self_values = self.__dict__
            for key in self_values.keys():
                if not getattr(self, key) == getattr(__o, key):
                    return False
            return True
    

    @classmethod
    def init_dict(cls, dict_values=dict):
        try:
            sheets_name = dict_values['sheets_name']
            sheets_name_done = dict_values['sheets_name_done']
            sheets_id = dict_values['sheets_id']
            sheets_range = dict_values['sheets_range']
            sheets_range_done = dict_values['sheets_range_done']
            try_number = dict_values['try_number']
            values_list = dict_values['values_list']
            return cls(sheets_name, sheets_name_done, sheets_id, sheets_range, sheets_range_done, try_number, values_list)
        except Exception as error:
            logger.error(f'Error in {error}')
    

@dataclass
class Product:
    code : any
    description : str
    type : str
    price_1 : float
    price_2 : float
    price_3 : float
    price_4 : float
    price_5 : float
    price_6 : float
    repos_cost : float


    def __eq__(self, __o: object) -> bool:
        if not isinstance(__o, self.__class__):
            return NotImplemented
        else:
            self_values = self.__dict__
            for key in self_values.keys():
                if not getattr(self, key) == getattr(__o, key):
                    return False
            return True


    @classmethod
    def init_dict(cls, dict_values=dict):
        try:
            code = dict_values['code']
            description = dict_values['description']
            type = dict_values['type']
            price_1 = dict_values['price_1']
            price_2 = dict_values['price_2']
            price_3 = dict_values['price_3']
            price_4 = dict_values['price_4']
            price_5 = dict_values['price_5']
            price_6 = dict_values['price_6']
            repos_cost = dict_values['repos_cost']
            return cls(code, description, type, price_1, price_2, price_3, price_4, price_5, price_6, repos_cost)
        except Exception as error:
            logger.error(f'Erro in {error}')