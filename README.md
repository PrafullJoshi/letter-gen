# Letter PDF Generator
## Lombok
- This project used lombok. Make sure IDE is setup with proper settings
- Refer https://www.baeldung.com/lombok-ide for more details

## Itext
- Free version of itest/poi is being used
- No commercial license required

## Swagger
- Has swagger end point configuration
- Once the server is spinned up access - http://localhost:7373/letter-gen/swagger-ui.html
- http://localhost:7373/letter-gen/swagger-ui.html#/Letter%20Generator/generateLettersUsingPUT is the Letter Generator API to be called

## Letter PDF
- API - http://localhost:7373/letter-gen/swagger-ui.html#/Letter%20Generator/generateLettersUsingPUT
- No need to pass any DTO in input Payload for now
