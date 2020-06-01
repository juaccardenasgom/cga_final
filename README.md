# Parametrizador
## Uso
Se utiliza el endpoint ```/api/create``` para la creación de archivos Excel para la parametrización a partir de un archivo JSON.
```
POST /api/create

body: JSON 
{
  "data": {
    "type#n": [ { 
      "piece#n": {
        "field#n": value
      }  
    } ]
  }
}

response: EXCEL FILE
```