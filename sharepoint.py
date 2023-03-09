import pandas as pd
import io
import openpyxl

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File


class Config():
    
    def set_parameters(self, params):
        '''
        Check conditions parameters for the process
        Attributes
        ----------
            url : dtype str
                Area sharepoint url.
                --> example = 'https://gerdaucld.sharepoint.com/sites/data.analytics'

            client_id : dtype str
                Website access email or website security credentials
                -> url to help = 'https://learn.microsoft.com/pt-br/sharepoint/dev/solution-guidance/security-apponly-azureacs'
                
            client_secret : dtype str
                Website access network password or website security credentials
                -> url to help = 'https://learn.microsoft.com/pt-br/sharepoint/dev/solution-guidance/security-apponly-azureacs'
               
        Returns
        ----------
            Dictionary with parameters
        '''
        
        self.params = {
            'url':None,
            'client_id':None,
            'client_secret':None
        }
        
        for arg in params:
            if arg in self.params:
                self.params[arg] = params[arg]
            else:
                raise ValueError(f'Parâmetro desconhecido: {arg}, esperado {[key for key in self.params]}')
        
        return self.set_sharepoint_conn()
        

    def set_sharepoint_conn(self):        
        self.__sharepoint_url = self.params.get('url')
        self.__sharepoint_client_id = self.params.get('client_id')
        self.__sharepoint_client_secret = self.params.get('client_secret')
        
    def get_sharepoint_url(self):
        return self.__sharepoint_url 

    def get_sharepoint_client_id(self):
        return self.__sharepoint_client_id

    def get_sharepoint_client_secret(self):
        return self.__sharepoint_client_secret 
    

class SharePoint(Config):
    
    def __init__(self, params: dict):
        self.set_parameters(params)
        
        self.CLIENT_ID = self.get_sharepoint_client_id()
        self.CLIENT_SECRET= self.get_sharepoint_client_secret()
        self.SHAREPOINT_URL = self.get_sharepoint_url()
        
        
    def auth(self):
        '''
            Website access authentication
            
        Attributes
        ----------
            
        Returns
        ----------
           Authentication for Client
        '''
        
        e_mail = True if len(self.CLIENT_ID.split('@gerdau.com')) > 1 else False
        ctx_auth = AuthenticationContext(self.SHAREPOINT_URL)
                
        if e_mail:
            ctx_auth.acquire_token_for_user(self.CLIENT_ID, self.CLIENT_SECRET)
        else:
            ctx_auth.acquire_token_for_app(self.CLIENT_ID, self.CLIENT_SECRET)

        ctx = ClientContext(self.SHAREPOINT_URL, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print("Authentication successful")
        print("Web site title: {0}".format(web.properties['Title']))
        
        return ctx
    
    
    def get_file(self, file_name: str, folder: str, **kwargs):
        '''
            Selection method for get file of the sharepoint as type Excel
            
        Attributes
        ----------
            file_name : dtype str
                Excel format file name
                --> example = 'file_name.xlsx'
            
            folder : dtype str
                Address where the file is found
                --> example = 'Shared Documents/folder'
            
            kwargs : Any parameters to use pandas.read_excel() or pandas.read_csv()
                        url:'https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html'
                        url:'https://pandas.pydata.org/docs/reference/api/pandas.read_csv.html'
   
        Returns
        ----------
           DataFrame with sharepoint file type Excel
        '''
        
        self.auth_site = self.auth()
        type_file = file_name.split('.')[-1]
        url = self.SHAREPOINT_URL.split('.com')[-1]
        folder_name = url + '/' + folder + "/" + file_name
        
        response = File.open_binary(self.auth_site, folder_name)

        #save data to BytesIO stream
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(response.content)
        bytes_file_obj.seek(0)

        if type_file == 'csv':
            encoding = kwargs.get('encoding')
            
            if encoding is None:
                df = pd.read_csv(bytes_file_obj, encoding='utf8', **kwargs)
            else:
                df = pd.read_csv(bytes_file_obj, **kwargs)
                             
        elif type_file in ['xls','xlsx','xlsm','xlsb']:  
            engine = kwargs.get('engine')
            
            if engine is None:
                df = pd.read_excel(bytes_file_obj, engine='openpyxl', **kwargs)
            else: 
                df = pd.read_excel(bytes_file_obj, **kwargs)
                
        else:
            raise ValueError(
                f'Arquivo do tipo .{type_file}, não suportando para extração do mesmo'
                )
            
        return df
    
    
    def get_list(self, ls_name: str):
        '''
            Selection method for get List of the sharepoint
            
        Attributes
        ----------
            ls_name : dtype str
                List name registered in sharepoint
                --> example: https://gerdaucld.sharepoint.com/sites/data.analytics/Lists/tb_repositorio/AllItems.aspx
                ---> ls_name: "tb_repositorio"
        
        Returns
        ----------
            DataFrame with sharepoint list
        '''
        
        self.auth_site = self.auth()
        
        lists = self.auth_site.web.lists
        result_lists = lists.get_by_title(ls_name)
        lists_items = result_lists.get_items()
        self.auth_site.load(lists_items)
        self.auth_site.execute_query()
        
        lists_data = []
        for idx, item in enumerate(lists_items):
            lists_data.append(item.properties)

        return pd.DataFrame(lists_data)