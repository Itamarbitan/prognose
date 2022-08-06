const XLSX = require('xlsx');
const path = require('path');
const POST = 3000;
const express = require('express');
const app = express();
app.set('view engine', 'ejs')
const staticWebsite = path.join(__dirname , './views')

app.use(express.static(staticWebsite))


bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: false }));




const exportUsersToExcel = (UserList, workSheetColumnName, workSheetName, filePath) => {

    const data = UserList.map(user => {
        return [user.weather , user.food, user.hour , user.positive_on_positive , user.negetive_on_positive, user.positive_on_negetive, user.negetive_on_negetive];
    });
    const workBook = XLSX.utils.book_new();
    const workSheetData = [
        workSheetColumnName,
        ...data
    ];
    const workSheet = XLSX.utils.aoa_to_sheet(workSheetData);
    // XLSX.utils.sheet_add_aoa(worksheet, [data], UserList);
    XLSX.utils.book_append_sheet(workBook, workSheet ,workSheetName)
    XLSX.writeFile(workBook, path.resolve(filePath));
    return true;

};





let UserList = [];

const workSheetColumnName = [
    'weather',
    'food',
    'hour',
    'positive_on_positive',
    'negetive_on_positive',
    'positive_on_negetive',
    'negetive_on_negetive'

];


const workSheetName = 'Users';
const filePath = './users.xlsx';


app.post('' , (req , res) => {
    console.log('hey');
    res.render('index');
    UserList = [{
        'weather' : `${req.body.weather}`,
        'food' : `${req.body.food}`,
        'hour' : `${req.body.hour}`,
        'positive_on_positive' : `${req.body.positive_on_positive}`,
        'negetive_on_positive' : `${req.body.negetive_on_positive}`,
        'positive_on_negetive' : `${req.body.positive_on_negetive}`,
        'negetive_on_negetive' : `${req.body.negetive_on_negetive}`
    
    }];
    console.log(UserList);


    exportUsersToExcel(UserList , workSheetColumnName , workSheetName , filePath);

});




app.get('' , (req , res) => {
    res.render('index')}
);


app.listen(POST , () => console.log('the app is listening to port 3000!!'));
