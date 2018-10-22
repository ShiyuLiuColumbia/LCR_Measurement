#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QtSerialPort/QSerialPort>
#include <QtSerialPort/QSerialPortInfo>
#include <QTimer>
#include <QTime>
#include <QDebug>
#include <ActiveQt/QAxObject>
#include <QFileDialog>
#include <QStandardItemModel>
#include <QMessageBox>
#include <windows.h>
#include <QElapsedTimer>
#include <exception>


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);


    foreach( const QSerialPortInfo &Info,QSerialPortInfo::availablePorts())//读取串口信息
        {
            //qDebug() << "portName    :"  << Info.portName();//调试时可以看的串口信息
            //qDebug() << "Description   :" << Info.description();
            //qDebug() << "Manufacturer:" << Info.manufacturer();

            QSerialPort serial;
            serial.setPort(Info);
            //if (ui->comboBox)
            if( serial.open( QIODevice::ReadWrite) )//如果串口是可以读写方式打开的
            {
                ui->comboBox->addItem( Info.portName() );//在comboBox那添加串口号
                ui->comboBox_2->addItem( Info.portName() );//在comboBox那添加串口号
                serial.close();//然后自动关闭等待人为开启（通过那个打开串口的PushButton）
            }
        }

    excel = new QAxObject(this);


    excel->setControl("Excel.Application");//连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
    excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示


    timer2 = new QTimer(this);
    connect( timer2, SIGNAL( timeout() ), this, SLOT( readMyCom2() ) );
    connect(my_serialPort1,SIGNAL(readyRead()),this,SLOT(readMyCom1()));
    connect(this,SIGNAL(measureready( QList<QList<QVariant> >)),this,SLOT(measure(QList<QList<QVariant> >)));
    connect(this,SIGNAL(LCRready()),this,SLOT(LCRmeasure()));
    connect(this,SIGNAL(measure2ready()),this,SLOT(measure2()));





    ui->tab->setEnabled(false);




}

MainWindow::~MainWindow()
{
    delete ui;
    if(openexcel){
    _workbook->dynamicCall("Close");
    }
    excel->dynamicCall("Quit()");//关闭excel
    delete excel;
}

void MainWindow::on_pushButton_clicked()

{
    QString path = QFileDialog::getOpenFileName(this, tr("读取excel文件"),".",tr("Excel Files(*.xlsx *.xls)"));
    qDebug()<<path;
    if(path.isEmpty()) {

        }
    else {

    QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合

    QAxObject *workbook = workbooks->querySubObject("Open(QString, QVariant)", path);
    QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
    QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
    QAxObject * usedrange = worksheet->querySubObject("UsedRange");
    QVariant var;
    var = usedrange->dynamicCall("Value");
    QAxObject * rows = usedrange->querySubObject("Rows");
    QAxObject * columns = usedrange->querySubObject("Columns");
    int row_start = usedrange->property("Row").toInt();  //获取起始行
    int column_start = usedrange->property("Column").toInt();  //获取起始列
    int row_count = rows->property("Count").toInt();  //获取行数
    int column_count = columns->property("Count").toInt();  //获取列数
    QStandardItemModel *model = new QStandardItemModel(row_count-1, column_count-1);


    res.clear();
    castVariant2ListListVariant(var,res);

    for(int i=row_start; i<=row_count; ++i)
    {
        for(int j=column_start+1; j<=column_count; ++j)
        {
            QVariant cell = res[i-1][j-1];

            QString cell_value = cell.toString();  //获取单元格内容
            //QAxObject *cell = worksheet->querySubObject("Cells(int,int)", i, j);

           // QString cell_value = cell->property("Value2").toString();  //获取单元格内容
            if( i == 1 )
                model->setHeaderData(j-2, Qt::Horizontal, cell_value);//将表的列名，放入model的列名中
            else
                model->setData(model->index(i-2, j-2, QModelIndex()), cell_value);

        }
    }

    qDebug() << "xls行数："<<row_count;
    qDebug() << "xls列数："<<column_count;
    workbooks->dynamicCall("Close()");
    ui->tableView->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);


    ui->tableView->setModel(model);//将model与view关联













    }
}

void MainWindow::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res)
{
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
        return;
    }
    const int rowCount = varRows.size();
    qDebug()<< rowCount;
    QVariantList rowData;
    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}




void MainWindow::measure(QList<QList<QVariant> > res)
{


       if(measure3){
        ui->pushButton_2->setEnabled(false);
        ui->pushButton_3->setEnabled(false);
        ui->pushButton_4->setEnabled(false);
        ui->pushButton_5->setEnabled(false);
       }

        int measure_line =  res[0].length()-1;
        int measure_column = res.length()-1;
        if(clicktime<=measure_column){
        QByteArray tmp;
        tmp[0] =12;
        qDebug()<<"once";
       // tmp[1] =2;
        my_serialPort1->write(tmp);

        QByteArray tmp1;
        tmp1[0] =0;
        QByteArray tmp2;
        tmp2[0] =1;
        QByteArray tmp3;
        tmp3[0] =2;
        QByteArray tmp4;
        tmp4[0] =3;


            for(int p=1;p<=measure_line;p++)
            {
        if (res[clicktime][p]=="激励")
        {qDebug()<<"激励";
            my_serialPort1->write(tmp2);

        }
        else if(res[clicktime][p]=="测量")
        {
           qDebug()<<"测量";
            my_serialPort1->write(tmp3);

        }

        else if(res[clicktime][p]=="接地")
        {
            qDebug()<<"接地";
            my_serialPort1->write(tmp1);
        }
        else if(res[clicktime][p]=="浮空")
        {
            qDebug()<<"浮空";
            my_serialPort1->write(tmp4);
        }
}

            clicktime++;
    }

    }





void MainWindow::readMyCom1()//读取缓冲的数据，每秒读一次
{


   if(measure3==true){
    requestData1.clear();    //清除缓冲区//
    requestData1 = my_serialPort1->readAll();//用requestData存储从串口那读取的数据
    qDebug() << requestData1;
    if (requestData1.toInt()==4){

        emit LCRready();

    }
   }
}

void MainWindow::LCRmeasure(){
    QByteArray tmp5;
    tmp5[0]=42;
    tmp5[1]=84;
    tmp5[2]=82;
    tmp5[3]=71;
    tmp5[4]=10;
    my_serialPort2->write(tmp5);

    timer2->setSingleShot(true);

    timer2->start(1000);//每秒读一次



}
void MainWindow::readMyCom2()//读取缓冲的数据，每秒读一次
{

    requestData2.clear();    //清除缓冲区//
    requestData2 = my_serialPort2->readAll();//用requestData存储从串口那读取的数据
    qDebug() << requestData2;
    if (requestData2!=""){
      //  qDebug()<<requestData2.toInt();
       savedata();

  int measure_column = res.length()-1;

if((savecount)!=measure_column){
        emit measureready(res);
    savecount++;

}

else{if(multimeasure==true){
    int a = ui->spinBox->text().toInt();
    if(measuretime!=a){

        measuretime++;
        emit measure2ready();
    }else{       
   QMessageBox::about(ui->tab_2,tr("提示"),tr("测量结束"));
   ui->pushButton_2->setEnabled(true);
   ui->pushButton_3->setEnabled(true);
   ui->pushButton_4->setEnabled(true);
   ui->pushButton_5->setEnabled(true);
    }}else{
        QMessageBox::about(ui->tab_2,tr("提示"),tr("测量结束"));
        ui->pushButton_2->setEnabled(true);
        ui->pushButton_3->setEnabled(true);
        ui->pushButton_4->setEnabled(true);
        ui->pushButton_5->setEnabled(true);
    }

    }
}


 }

void MainWindow::savedata(){


    QString s1;
    QStringList s2;
    QString s3;
    QString s4;
    s1 = requestData2.data();
   // s1="+0,+6.07790E-08,+1.50825E-03\n";
    s2 = s1.split(",");
    s3 = s2.at(1);
    s4 = s2.at(2);

    QString s5 = s4;
   // s3=s3.remove(0, 1);
   // s4=s4.remove(0, 1);
    s4 = s4.left(s4.length()-1);
    double date1 = s3.toDouble();
    double date2 = s4.toDouble();
    ui->textBrowser_2->insertPlainText(s3+"  "+s5);
    qDebug()<<date1<<date2;
    int measure_column = res.length()-1;
    write2excel(savecount+1,column_start+column_count,date1);
    write2excel(savecount+1,column_start+column_count+1,date2);
    ui->progressBar->setValue((savecount+1)+measure_column*(measuretime-1));





}

void MainWindow::on_pushButton_7_clicked()
{

         bool tmp1 = false;
         bool tmp2 = false;



         my_serialPort1->setPortName( ui->comboBox->currentText() );
        tmp1= my_serialPort1->open( QIODevice::ReadWrite );
         qDebug() << ui->comboBox->currentText();
         my_serialPort1->setBaudRate(9600);//波特率
         my_serialPort1->setDataBits( QSerialPort::Data8 );//数据字节，8字节
         my_serialPort1->setParity( QSerialPort::NoParity );//校验，无
         my_serialPort1->setFlowControl( QSerialPort::NoFlowControl );//数据流控制,无

         my_serialPort2->setPortName( ui->comboBox_2->currentText() );
         tmp2=my_serialPort2->open( QIODevice::ReadWrite );
         qDebug() << ui->comboBox_2->currentText();
         my_serialPort2->setBaudRate(9600);//波特率
         my_serialPort2->setDataBits( QSerialPort::Data8 );//数据字节，8字节
         my_serialPort2->setParity( QSerialPort::NoParity );//校验，无
         my_serialPort2->setFlowControl( QSerialPort::NoFlowControl );//数据流控制,无
         my_serialPort2->setStopBits( QSerialPort::OneStop );//一位停止位    newthread.my_serialPort1->setStopBits( QSerialPort::OneStop );//一位停止位
         if(tmp1&&tmp2&&((ui->comboBox->currentText())!=(ui->comboBox_2->currentText())) )
         { QMessageBox::about(ui->tab_2,tr("提示信息"),tr("串口打开成功"));
             ui->tab->setEnabled(true);

          }
         else
         {QMessageBox::critical(ui->tab_2,tr("错误信息 "),tr("串口打开错误"));
         }




}


void MainWindow::on_pushButton_8_clicked()
{          my_serialPort1->close();
           my_serialPort2->close();

            ui->comboBox->clear();
            ui->comboBox->insertItems(0, QStringList()
             << QString()
            );
            ui->comboBox_2->clear();
            ui->comboBox_2->insertItems(0, QStringList()
             << QString());
    foreach( const QSerialPortInfo &Info,QSerialPortInfo::availablePorts())//读取串口信息
        {
            //qDebug() << "portName    :"  << Info.portName();//调试时可以看的串口信息
            //qDebug() << "Description   :" << Info.description();
            //qDebug() << "Manufacturer:" << Info.manufacturer();

            QSerialPort serial;
            serial.setPort(Info);
            if (ui->comboBox)
            if( serial.open( QIODevice::ReadWrite) )//如果串口是可以读写方式打开的
            {
                ui->comboBox->addItem( Info.portName() );//在comboBox那添加串口号
                ui->comboBox_2->addItem( Info.portName() );//在comboBox那添加串口号
                serial.close();//然后自动关闭等待人为开启（通过那个打开串口的PushButton）
            }
        }

}

void MainWindow::on_pushButton_2_clicked()
{  openexcel =true;
   measuretime =1;
    clicktime=1;
    savecount =1;
    multimeasure = false;
    measure3 =true;

    if(ui->textBrowser->toPlainText().isEmpty()){
        QMessageBox::critical(ui->tab_2,tr("错误信息"),tr("保存路径为空"));
    }
    else{
        if(res.length() == 0)
        {
            QMessageBox::critical(ui->tab_2,tr("错误信息 "),tr("EXCEL没有数据"));
        }
        else{
    QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
   _workbook = workbooks->querySubObject("Open(QString, QVariant)", fileName);
    worksheets = _workbook->querySubObject("Sheets");//获取工作表集合
    worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
     usedrange = worksheet->querySubObject("UsedRange");
     columns = usedrange->querySubObject("Columns");
     column_start = usedrange->property("Column").toInt();  //获取起始列
     column_count = columns->property("Count").toInt();  //获取列数

     QAxObject *cell = worksheet->querySubObject("Cells(int, int)", 1, column_start+column_count);//等同于上一句
     cell->dynamicCall("SetValue(const QVariant&)",QVariant(tr("空场标定")));//存储一个double 数据到 excel 的单元格中
      _workbook->dynamicCall("Save()");
      int measure_column = res.length()-1;
      ui->progressBar->setRange(1,measure_column);
      ui->progressBar->setValue(1);

    emit measureready(res);}
}


}

void MainWindow::on_pushButton_3_clicked()
{   openexcel =true;
    measuretime =1;
    clicktime=1;
       savecount =1;
          multimeasure = false;
          measure3 =true;
          if(ui->textBrowser->toPlainText().isEmpty()){
              QMessageBox::critical(ui->tab_2,tr("错误信息"),tr("保存路径为空"));
          }
          else{
              if(res.length() == 0)
              {
                  QMessageBox::critical(ui->tab_2,tr("错误信息 "),tr("EXCEL没有数据"));
              }
              else{
          QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
         _workbook = workbooks->querySubObject("Open(QString, QVariant)", fileName);
          worksheets = _workbook->querySubObject("Sheets");//获取工作表集合
          worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
           usedrange = worksheet->querySubObject("UsedRange");
           columns = usedrange->querySubObject("Columns");
           column_start = usedrange->property("Column").toInt();  //获取起始列
           column_count = columns->property("Count").toInt();  //获取列数
              QAxObject *cell = worksheet->querySubObject("Cells(int, int)", 1, column_start+column_count);//等同于上一句
              cell->dynamicCall("SetValue(const QVariant&)",QVariant(tr("满场标定")));//存储一个double 数据到 excel 的单元格中
               _workbook->dynamicCall("Save()");
               int measure_column = res.length()-1;
               ui->progressBar->setRange(1,measure_column);
               ui->progressBar->setValue(1);

       emit measureready(res);}
          }
}

void MainWindow::on_pushButton_4_clicked()
{  openexcel =true;
    measuretime =1;
    clicktime=1;
    savecount =1;
    multimeasure = false;
    measure3 =true;
    if(ui->textBrowser->toPlainText().isEmpty()){
        QMessageBox::critical(ui->tab_2,tr("错误信息"),tr("保存路径为空"));
    }
    else
    {    if(res.length() == 0)
        {
            QMessageBox::critical(ui->tab_2,tr("错误信息 "),tr("EXCEL没有数据"));
        }
        else{
    QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
   _workbook = workbooks->querySubObject("Open(QString, QVariant)", fileName);
    worksheets = _workbook->querySubObject("Sheets");//获取工作表集合
    worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
     usedrange = worksheet->querySubObject("UsedRange");
     columns = usedrange->querySubObject("Columns");
     column_start = usedrange->property("Column").toInt();  //获取起始列
     column_count = columns->property("Count").toInt();  //获取列数
     QAxObject *cell = worksheet->querySubObject("Cells(int, int)", 1, column_start+column_count);//等同于上一句
     cell->dynamicCall("SetValue(const QVariant&)",QVariant(tr("单次测量")));//存储一个double 数据到 excel 的单元格中
      _workbook->dynamicCall("Save()");
      int measure_column = res.length()-1;
      ui->progressBar->setRange(1,measure_column);
      ui->progressBar->setValue(1);


    emit measureready(res);}
    }
}

/*void MainWindow::on_pushButton_9_clicked()
{
    QString filepath = QFileDialog::getSaveFileName(NULL,"Save File",".","Excel File (*.xls)");

    if(filepath.isEmpty()) {

        }
    else {
        filepath.replace("/","\\");  //这一步很重要，c:/123.xls保存失败，c:\123.xls保存成功！
        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
        workbooks->dynamicCall("Add");//新建一个工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
        workbook->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",filepath,56,QString(""),QString(""),false,false);

      ui->textBrowser->clear();
     ui->textBrowser->insertPlainText(filepath);
     fileName = filepath;
    }


}*/

void MainWindow::write2excel(int row,int column,double value)
{
    worksheets = _workbook->querySubObject("Sheets");//获取工作表集合
       worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
    QAxObject *cell = worksheet->querySubObject("Cells(int, int)", row, column);//等同于上一句
    cell->dynamicCall("SetValue(const QVariant&)",QVariant(value));//存储一个double 数据到 excel 的单元格中
     _workbook->dynamicCall("Save()");


}

void MainWindow::on_pushButton_10_clicked()
{
     QString filepath = QFileDialog::getOpenFileName(this, tr("读取excel文件"),".",tr("Excel Files(*.xlsx *.xls)"));
    if(filepath.isEmpty()) {

    }else{


    ui->textBrowser->clear();
    ui->textBrowser->insertPlainText(filepath);
    fileName = filepath;

    }


}

void MainWindow::on_pushButton_5_clicked()
{  int measure_column = res.length()-1;
    ui->progressBar->setRange(1,measure_column*(ui->spinBox->text().toInt()));
    ui->progressBar->setValue(1);
    measuretime = 1;
    multimeasure = true;

emit measure2ready();


}


void MainWindow::measure2(){

    clicktime=1;
    savecount =1;
    measure3 =true;
    if(ui->textBrowser->toPlainText().isEmpty()){
        QMessageBox::critical(ui->tab_2,tr("错误信息"),tr("保存路径为空"));}
    else{
         if(res.length() == 0)
        {
            QMessageBox::critical(ui->tab_2,tr("错误信息 "),tr("EXCEL没有数据"));
        }
        else{
    QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
   _workbook = workbooks->querySubObject("Open(QString, QVariant)", fileName);
    worksheets = _workbook->querySubObject("Sheets");//获取工作表集合
    worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
     usedrange = worksheet->querySubObject("UsedRange");
     columns = usedrange->querySubObject("Columns");
     column_start = usedrange->property("Column").toInt();  //获取起始列
     column_count = columns->property("Count").toInt();  //获取列数
     QString s1=tr("连续测量第") ;
     QString s2=QString::number(measuretime, 10);
     QString s4=tr("次") ;
     QString s3 = s1+s2+s4;
     qDebug()<<s3;
             QAxObject *cell = worksheet->querySubObject("Cells(int, int)", 1, column_start+column_count);//等同于上一句
     cell->dynamicCall("SetValue(const QVariant&)",QVariant(s3));//存储一个double 数据到 excel 的单元格中
      _workbook->dynamicCall("Save()");

    emit measureready(res);}
}
}

void MainWindow::on_pushButton_6_clicked()
{   clicktime = 1;
    measure3 =false;
    if(res.length() == 0)
           {
               QMessageBox::critical(ui->tab_2,tr("错误信息"),tr("EXCEL没有数据"));
           }
           else{
       emit measureready(res);
        QMessageBox::about(ui->tab_2,tr("提示"),tr("开始短路/开路标定，电极状态为Excel第一行"));
       }

}
