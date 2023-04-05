# Conector Python do SharePoint
Este pacote visa fornecer uma maneira de conectar-se a sites do SharePoint e recuperar arquivos Excel ou listas como DataFrame pandas.

# Requisitos
pandas <br>
openpyxl <br>
office365_python_sdk <br>

# Instalação
![image](https://user-images.githubusercontent.com/78990428/230177524-432dfc79-cf56-4192-8fb4-46e028526dae.png)

# Como usar
## Autenticação
Para se conectar a um site do SharePoint, você precisa criar uma `Config` classe com os <br>
parâmetros `url`, `client_id` e `client_secret`. Esses parâmetros podem ser <br>
encontrados em seu site do SharePoint. Depois de criar a `Config` classe, você pode usá-la <br>
para criar uma SharePoint instância de classe.

![image](https://user-images.githubusercontent.com/78990428/230179098-54db0b48-3262-482f-ad90-4b875d49a91f.png)

# Recuperando arquivos do Excel
Para recuperar um arquivo do Excel, use o `get_file` método. Você precisa passar os <br>
parâmetros `folder` e `file_name`, que são o caminho para o arquivo em seu site do <br>
SharePoint. Você também pode passar quaisquer parâmetros adicionais que `read_excel` a <br>
função do pandas aceite.

![image](https://user-images.githubusercontent.com/78990428/230179199-bca24e84-e17d-40cb-aa0e-ac9d64527804.png)

# Recuperando listas do SharePoint
Para recuperar uma lista do SharePoint, use o `get_list` método. Você precisa passar o <br>
`ls_name` parâmetro, que é o nome da lista no seu site do SharePoint.

![image](https://user-images.githubusercontent.com/78990428/230179266-dbcf5203-e5c7-441a-b18b-6ad532aa3952.png)
