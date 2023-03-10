import pandas as pd
from io import BytesIO
import openpyxl

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File


class Config:
    
    def set_parameters(self, params) -> dict:
        '''
        Check conditions parameters for the process
            Attributes
            ----------
                url : dtype str
                    Area sharepoint url.
                    --> example = 'https://sharepoint.com/sites/dataanalytics'

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
                raise ValueError(f'Unknown parameter: {arg}, expected {[key for key in self.params]}')
        
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
    
    def __init__(self, params:dict):
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
        
        e_mail = True if len([tx for tx in ['@', '.com','.com.'] if tx in self.CLIENT_ID]) >= 1 else False
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
    
    
    def get_file(self, 
            folder:str,
            file_name:str, 
            format_name:str, 
            **kwargs
        ) -> pd.DataFrame:
        '''
            Selection method for get file of the sharepoint as type Excel
            
        Attributes
        ----------
            file_name : dtype str
                Excel file name
                --> example = 'file_name.xlsx'
            
            format_name : dtype str
                Excel format name
                --> example = 'xls','xlsx','xlsm','xlsb'
            
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
        url = self.SHAREPOINT_URL.split('.com')[-1]
        folder_name = f'{url}/{folder}/{file_name}.{format_name}'
        
        response = File.open_binary(self.auth_site, folder_name)

        # Save data to BytesIO stream
        bytes_file_obj = BytesIO()
        bytes_file_obj.write(response.content)
        bytes_file_obj.seek(0)

        if format_name == 'csv':
            kwargs.update(encoding='utf8')
            
            return pd.read_csv(bytes_file_obj, **kwargs)
                             
        elif format_name in ['xls','xlsx','xlsm','xlsb']:
            kwargs.update(engine='openpyxl') 
            
            return pd.read_excel(bytes_file_obj, **kwargs)
                
        else:
            raise ValueError(
                f'File of type {format_name}, not supporting for extracting it'
                )
    
    
    def get_list(self, ls_name:str) -> pd.DataFrame:
        '''
        Selection method for get List of the sharepoint    
            Attributes
            ----------
                ls_name : dtype str
                    List name registered in sharepoint
                    --> example: https://sharepoint.com/sites/dataanalytics/Lists/tb_analytics
                    ---> ls_name: "tb_analytics"

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
        for _, item in enumerate(lists_items):
            lists_data.append(item.properties)

        return pd.DataFrame(lists_data)
