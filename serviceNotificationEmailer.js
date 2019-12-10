// ----- Set up -----
const nodemailer = require('nodemailer');
const sql = require('mssql');
const fs = require('fs');

let days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const config = {
  user: 'sa',
  password: 'password',
  server: 'svr',
  database: 'db'
}

// create reusable transporter object using the default SMTP transport
let transporter = nodemailer.createTransport({
  host: 'smtp.office365.com',
  port: 587,
  secure: false, // true for 465, false for other ports
  auth: {
    user: 'test', // generated ethereal user
    pass: 'test' // generated ethereal password
  }
});

let mainPool = new sql.ConnectionPool(config);
let mainPoolConnect = mainPool.connect();


mainPool.on("error", err => {
  console.log(err);
});

async function wrapedSendMail(mailOptions){
    return new Promise((resolve,reject)=>{
      transporter.sendMail(mailOptions, function(error, info){
        if (error) {
            console.log("error is "+error);
           resolve(false); // or use rejcet(false) but then you will have to handle errors
        }
        else {
           console.log('Email sent: ' + info.response);
           resolve(true);
        }
     });
   })
}
// ----- Get data from SQL server -----
//Open SQL connection
async function EmailServiceNotifications (){
  await mainPoolConnect; //checks if pool is connected
  try {
    console.log('Connection Established');
      let req = mainPool.request();
      let result = await req.query("Select top 30 mrwohst.id_pk,mrwohst.cust_id, mrwohst.loc_id, mrwohst.sys_id,SYS_TP_DSC, mrsls.sls_id, mrsls.f_name+' '+mrsls.l_name as salesperson,mrsls.e_mail_adr,mrsls.phone_no, svc_date, conf_date, rc_number, ifileloc, mrcust.email_1 as cust_email_1,mrcust.email_2 as  cust_email_2, mrcust.email_3 as cust_email_3, noemail, mrcust.atn, mrcust.[name], email as loc_email, mrloc.email_2 as loc_email_2, mrloc.EMAIL_3 as loc_email_3, enotify, mrloc.name as loc_Name,mrloc.adr1,mrloc.city, mrloc.[state], mrloc.zip From mrwohst left join mrcust on mrwohst.cust_id = mrcust.CUST_ID and mrcust.IsDeleted = 0 left join mrloc on mrloc.cust_id + mrloc.loc_id = mrwohst.cust_id + mrwohst.loc_id and mrloc.isdeleted = 0 left join mrsls on mrsls.SLS_ID = mrwohst.SLS_ID and mrsls.isdeleted = 0 left join mrsys on mrwohst.CUST_ID+mrwohst.LOC_ID+MRWOHST.SYS_ID = mrsys.CUST_ID+mrsys.LOC_ID+mrsys.SYS_ID and mrsys.IsDeleted = 0 left join mrsystp on mrsystp.SYS_TYPE = mrsys.SYS_TYPE and mrsystp.IsDeleted = 0 WHERE mrwohst.IsDeleted=0 And Status='1' And (ISNULL(email_date,'')=''  or year(email_date) = 1900) And conf_date>=GetDate()-14 Order By cust_id, loc_id, sys_id");
          if (result.rowsAffected > 0) {
            for (let i = 0; i < result.recordset.length; i++) {



              let file = `\\\\metro-nyc-web\\${result.recordset[i].cust_id.substring(3,4) == 'B'? "wohistory":"routecards"}${result.recordset[i].ifileloc}`;

              if (file.includes('norc.tif')) {
                file = '';
              }

              if (!fs.existsSync(file)) {
                file = '';
              }
              let htmlAttachmentMsg = '';

              if (file != ''){
                  htmlAttachmentMsg = `Attached, Please find a copy of your service report <br /><br />`;
              }

              htmlAttachmentMsg += `If you have any questions about this service visit, please feel free to contact me directly.<br /><br />`;

              let htmlBody = `From: The Metro Group Inc. ${result.recordset[i].cust_id.substring(3,1) == 'B'? "- Water Chemicals Division":""} <br />
                              Customer: ${result.recordset[i].name}<br />
                              Reference ID: ${result.recordset[i].cust_id+'-'+result.recordset[i].loc_id+'-'+result.recordset[i].sys_id+'-'+result.recordset[i].rc_number}<br /><br />
                              On ${days[result.recordset[i].svc_date.getDay()]} our service technician performed water treatment services at
                              ${result.recordset[i].adr1+' '+result.recordset[i].city+', '+result.recordset[i].state+' '+result.recordset[i].zip} <br /><br />
                              System ${result.recordset[i].sys_id}, ${result.recordset[i].SYS_TP_DSC}<br /><br />
                              ${htmlAttachmentMsg}
                              Sincerely,<br />
                              ${result.recordset[i].salesperson} <br />
                              Account Manager <br />
                              ${result.recordset[i].e_mail_adr}<br />
                              ${result.recordset[i].phone_no ?'Tel: '+result.recordset[i].phone_no : '' }
                              `;

              let emails = '';

              if (result.recordset[i].cust_email_1 != '' || result.recordset[i].cust_email_1 != null) {
                emails = result.recordset[i].cust_email_1;
              }
			  if (result.recordset[i].cust_email_2 != '' || result.recordset[i].cust_email_2 != null) {
                emails +=  emails != '' ? `, ${result.recordset[i].cust_email_2}`: result.recordset[i].cust_email_2;
              }
			  if (result.recordset[i].cust_email_3 != '' || result.recordset[i].cust_email_3 != null) {
                emails +=  emails != '' ? `, ${result.recordset[i].cust_email_3}`: result.recordset[i].cust_email_3;
              }
			  if (result.recordset[i].loc_email != '' || result.recordset[i].loc_email != null) {
                emails +=  emails != '' ? `, ${result.recordset[i].loc_email}`: result.recordset[i].loc_email;
              }
			  if (result.recordset[i].loc_email_2 != '' || result.recordset[i].loc_email_2 != null) {
                emails +=  emails != '' ? `, ${result.recordset[i].loc_email_2}`: result.recordset[i].loc_email_2;
              }
			  if (result.recordset[i].loc_email_3 != '' || result.recordset[i].loc_email_3 != null) {
                emails +=  emails != '' ? `, ${result.recordset[i].loc_email_3}`: result.recordset[i].loc_email_3;
              }
              console.log(emails);

               let mailOptions = {
                from: '"The Metro Group Inc." <auto-mail@metrogroupinc.com>', // sender address
                to: emails, // list of receivers
                replyTo: result.recordset[i].e_mail_adr,
                //to: 'smuratov@metrogroupinc.com',
                subject: `Automated Service Notification [${result.recordset[i].cust_id}-${result.recordset[i].loc_id}]`, // Subject line
                cc: '',
                bcc:'pdflog@metrogroupinc.com',
                //cc: '',
                html: htmlBody // html body
              };

              if (file != '') {
                  console.log('File Exists');
                  mailOptions.attachments = [{
                    filename: 'Service Report.tif',
                    path: file,
                    contentType: 'image/tiff'
                  }]
              }

              let today = new Date();
              let dd = today.getDate();
              let mm = today.getMonth() + 1;
              let yyyy = today.getFullYear();

              sendmail = async()=>{
                  let resp = await wrapedSendMail(mailOptions);
                  return resp;
              }

              let sendMailResp = await sendmail();
              if (sendMailResp){
                let req = mainPool.request();
                let insertResponse = await req
                  .input('sendname', sql.VarChar, 'Building / Property Manger')
                  .input('emailaddr', sql.VarChar, emails)
                  .input('ccaddr', sql.VarChar, result.recordset[i].e_mail_adr)
                  .input('msgtext', sql.VarChar, htmlBody)
                  .input('msubject', sql.VarChar, `Automated Service Notification [${result.recordset[i].cust_id}-${result.recordset[i].loc_id}]`)
                  .input('imgpathfil', sql.VarChar, !result.recordset[i].ifileloc.includes('norc.tif') ? result.recordset[i].ifileloc : '')
                  .input('sentflag', sql.Bit, 1)
                  .input('senttime', sql.DateTime, new Date())
                  .input('cstamp', sql.VarChar, `AUTOSERV${yyyy}${mm}${dd}`)
                  .query('INSERT INTO mservice (sendname, emailaddr, ccaddr, msgtext, msubject,imgpathfil,sentflag,senttime,cstamp) VALUES (@sendname, @emailaddr, @ccaddr, @msgtext, @msubject,@imgpathfil,@sentflag,@senttime,@cstamp)');
                  if (insertResponse.rowsAffected > 0) {
                      console.log('MSERVICE was added succesfully');
                       let updateMrwohst = await req
                         .input('date', sql.DateTime, new Date())
                         .query(`UPDATE mrwohst SET email_date = @date where id_pk = '${result.recordset[i].id_pk}'`);
                           if (updateMrwohst.rowsAffected > 0) {
                             console.log('Work Order History Updated');
                           }
                    }
              }
            }
            //Close connection to DB
            mainPool.close();
          } else {
            mainPool.close();
          }
    }
    catch(err){
      console.log(err.message);
	  process.exit;
      return err;
    }
}

EmailServiceNotifications();
