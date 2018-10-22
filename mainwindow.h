#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/QAxObject>
#include <QtSerialPort/QSerialPort>
#include <QtSerialPort/QSerialPortInfo>
#include <QThread>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res);
    void write2excel(int row,int column,double value);
private slots:
    void measure(QList<QList<QVariant> > res);
    void measure2();
    void on_pushButton_clicked();

    void on_pushButton_7_clicked();
    //void checkport();

    void on_pushButton_8_clicked();

    void on_pushButton_2_clicked();
    void readMyCom1();//用于读取数据
    void readMyCom2();//用于读取数据
    void on_pushButton_3_clicked();
    void LCRmeasure();

    void on_pushButton_4_clicked();

    //void on_pushButton_9_clicked();
    void savedata();

    void on_pushButton_10_clicked();

    void on_pushButton_5_clicked();

    void on_pushButton_6_clicked();

signals:
    void ready();
    void measureready( QList<QList<QVariant> >);
    void LCRready();
    void measure2ready();

private:
    Ui::MainWindow *ui;
    QList<QList<QVariant> > res;
    QSerialPort *my_serialPort1= new QSerialPort(this);//(实例化一个指向串口的指针，可以用于访问串口)
    QSerialPort *my_serialPort2 = new QSerialPort(this);//(实例化一个指向串口的指针，可以用于访问串口)


        QByteArray requestData1;//（用于存储从串口那读取的数据）
        QByteArray requestData2;//（用于存储从串口那读取的数据）
        int clicktime ;
        QTimer *timer2;
        QAxObject *excel;
        QString fileName;
        QAxObject *_workbook;
        int column_start = 0;
        int column_count = 0;
        QAxObject *worksheets;
        QAxObject *worksheet;
        QAxObject * usedrange;
        QAxObject * columns;
        int savecount;
        int measuretime=1;
        bool multimeasure ;
        bool measure3;
        bool openexcel = false;

        
};



#endif // MAINWINDOW_H
