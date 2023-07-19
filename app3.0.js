const {Client} = require('whatsapp-web.js');
const qrcode = require ('qrcode-terminal');
//let PE = [{}] //Arreglo para almacenar los mensajes pendientes

const client = new Client();
let IndicadorMasivo;

let REG = []; // Arreglo para almacenar los números registrados
let N_REG = []; // Arreglo para almacenar los números no registrados

const path = require('path');
const { strict } = require('assert');


//Definicion de la ruta donde se encuentran el archivo para los envios masivos
const route = path.dirname('z:/Privada/UDI/C1_CICLOS/SUBIDA MASIVA.xlsx' + path.join('/SUBIDA MASIVA.xlsx'))
console.log (route);

//Definicion de la ruta donde se encuentran el archivo para los envios masivos
const route2 = path.dirname('D:/NewTool/Archivos/SUBIDA MASIVA.xlsx' + path.join('/SUBIDA MASIVA.xlsx'))
console.log (route);


client.initialize();

//Metodo para generar el codigo qr
client.on ('qr', (qr)=>{
    console.log ('QR:',qr)
    qrcode.generate(qr,{small:true})
});

//Cuando se logra la conexion entrara en el siguiente bloque de codigo.
client.on ('ready', ()=>{
    console.log('Nueva sesion creada')

    client.on ('message_create', async (Sended) =>{
        
        
        //console.log ('sended:',Sended.body + "\n"+"Enviado por mi:" + Sended.fromMe)

        if (Sended.fromMe === false){
            IndicadorMasivo= false;
        } 
        

        if (Sended.body === "!CANCELAR"){
            IndicadorMasivo=false;
                for (let A of ENV){
                    ENV.splice(0, ENV.length);
                }
                for (let D of NR){
                    NR.splice(0, NR.length);
                }
                console.log("####-Limpiando caché-####")
                console.log (ENV, NR);
        }

        if (!Sended.fromMe){
            // crea un nuevo objeto `Date`
            let fecha_actual = new Date();
            // obtener la fecha y la hora
            let now = fecha_actual.toLocaleString();
            crearArchivoExcel(Sended.number, now, Sended.body)
        }



        //!MASIVO es la palabra clave para el empezar el envio
        if (Sended.body === '!MASIVO-C1'){
            const ENV = [{}]; //Arreglo para almacenar los mensajes enviados
            const NR = [{}]; //Arreglo para almacenar los Numeros no registrados
            const readerS = require ('xlsx')     
             IndicadorMasivo = true
            
             //Se definen las las variables para leer el archivo excel
            const dat = readerS.readFile(route)
            let sheet_name_list= dat.SheetNames
            let xlData = readerS.utils.sheet_to_json(dat.Sheets[sheet_name_list[0]])
            //console.log (xlData)


            //Se recorre el arreglo que tiene la informacion del excel
            console.log("Tamaño inicial del arreglo de mensajes enviados: ")
            console.log(ENV.length)
            console.log("Listado de números a verificar:")
                let j = 0;
                console.log(xlData.length)
                for (let D of xlData){
                    await delay(2000)
                    console.log(D.NUMERO)
                    let number = '506' + D.NUMERO
                    number = number.split(" ").join("")
                    let msgg = D.MENSAJE;
                    let chatid =  number + "@c.us"
                    let IDMSJ = D.IDMSG
                    let now = new Date()
                    let fecha = now.toLocaleDateString()
                    j = j+1;                   
                    if(isRegistered) {
                        client.sendMessage(chatid, msgg);
                        
                        //Si se envia el mensaje, se guarda dentro del arreglo de Enviados
                        ENV.push({IDMS:IDMSJ, Mensaje:msgg, Numero:number, Envio: 'Enviado: ', Hora: fecha,});
                    }else{
                        //Si NO se envia el mensaje, se guarda dentro del arreglo de No Registrados
                        NR.push({IDMS:IDMSJ, Mensaje:msgg, Numero:number, Envio: 'No Enviado: ', Hora: fecha,});
                    }  
                }
                if (j === xlData.length) {
                    generarReporte();
                }
                //console.log("REGISTROS DEL ARREGLO ENV")
                //console.log(ENV)
                //console.log("REGISTROS DEL ARREGLO NR")
                //console.log(NR)
        }
    });//fin de la funcion client.on
});









function generarReporte(){
    setTimeout(() => {
        if (ENV.length > 1) {
            //Despues de recorrer el arreglo con la totalidad de los masivos se genera el informe de envio
            let reader = require ('xlsx')
            let exc = reader.utils.book_new()
            let workSheet1 = reader.utils.json_to_sheet(NR)
            let workSheet2 = reader.utils.json_to_sheet(ENV)
            reader.utils.book_append_sheet(exc, workSheet1, 'NO REGISTRADO EN WHATSAPP')
            reader.utils.book_append_sheet(exc, workSheet2, 'ENVIADOS')
            reader.writeFile(exc, 'z:/Privada/UDI/C1_CICLOS/REPORTE ENVÍO.xlsx');               
            for (let A of ENV){
                ENV.splice(0, ENV.length);
            }
            for (let D of NR){
                NR.splice(0, NR.length);
            }
            // crea un nuevo objeto `Date`
            let fecha_actual = new Date();
            // obtener la fecha y la hora
            let now = fecha_actual.toLocaleString();
            console.log("###################################################")
            console.log("####-Informe Generado Ciclos-####")
            console.log("fecha de creacion de informe: " +now)
            console.log("###################################################"+"\n\n")
        }else{
            // crea un nuevo objeto `Date`
            let fecha_actual = new Date();
            // obtener la fecha y la hora
            let now = fecha_actual.toLocaleString();
            console.log("###################################################")
            console.log("####-No se realizo ningún envío-####") 
            console.log("No se genero el informe: " +now)
            console.log("###################################################"+"\n\n") 
        }

    
    }, 8000);//se indica que la funcion se ejecutará despues de 3 segundos
}

//Función para ingresar pausa en la ejecucion del código
function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}


//Función para verificar el archivo Excel existe
function existeArchivoExcel(numeroCelular) {
    const nombreArchivo = `MSJ-${numeroCelular}.xlsx`;
    return fs.existsSync(nombreArchivo);
}


//Función para crear el archivo Excel
function crearArchivoExcel(numeroCelular, fecha, mensaje) {
    let ExcelJS = require ('exceljs')
    const fs = require ('fs');
    const nombreArchivo = `MSJ-${numeroCelular}.xlsx`;
    const workbook = new ExcelJS.Workbook();

    if (!fs.existsSync(nombreArchivo)) {
        const worksheet = workbook.addWorksheet('Datos'); // Nombre de la hoja de cálculo
        // Si el archivo no existe, podemos agregar encabezados o información adicional a la hoja de cálculo.
        worksheet.getCell('A1').value = 'Fecha';
        worksheet.getCell('B1').value = 'Mensaje';
    } else {
        workbook.xlsx.readFile(nombreArchivo)
        .catch((err) => {
        console.error('Error al leer el archivo Excel:', err);
        });
    }

    const worksheet = workbook.getWorksheet('Datos');

    // Aquí agregamos la nueva fila con la fecha y mensaje.
    worksheet.addRow([fecha, mensaje]);
  
    return workbook.xlsx.writeFile(nombreArchivo)
      .then(() => {
        console.log(`Se ha ${fs.existsSync(nombreArchivo) ? 'actualizado' : 'creado'} el archivo Excel: ${nombreArchivo}`);
      })
      .catch((err) => {
        console.error('Error al escribir en el archivo Excel:', err);
      });
}