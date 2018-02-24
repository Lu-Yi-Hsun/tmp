package main

import (

    "./lib"

    "fmt"
    "github.com/360EntSecGroup-Skylar/excelize"

    "strconv"

)


//var excel_col=[...]string{"A", "B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"}

func main() {


    fmt.Println("轉換中.....")

 //  s,err:=lib.Get_day_info("./excel/Nail library歌劇院店20172月營業報.xlsx",2,10)
//fmt.Println(s,err)
    xlsx, err := excelize.OpenFile("./設定檔案/設定檔案.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }
    //拿到檔名
   now_filename:=xlsx.GetCellValue("台中","B84")
   pass_filename:=xlsx.GetCellValue("台中","F84")
    year,month,day:=lib.Get_month_day(xlsx.GetCellValue("台中","B86"))



  //開啟想要時間的檔案
    now, err := excelize.OpenFile("./excel/"+now_filename)
    if err != nil {
        fmt.Println(err)
        return
    }
    //開啟過去時間的檔案
  //  pass, err := excelize.OpenFile("./excel/"+pass_filename)
    if err != nil {
        fmt.Println(err)
        return
    }

    work_day,err:=lib.Get_work_day("./excel/"+now_filename)
    if err!=nil{
        fmt.Println(err)
    }


    //目標設定
    sales_set_all :=xlsx.GetCellValue("台中","B79")
    //商品
    commodity__set_all  :=xlsx.GetCellValue("台中","B75")
    //技術回收
    tech_set_all  :=xlsx.GetCellValue("台中","B73")
    //禮券
    gift_set_all  :=xlsx.GetCellValue("台中","B76")
    //沙龍
    salon_set_all  :=xlsx.GetCellValue("台中","B77")
    //服務單pay
    serve_set_all  :=xlsx.GetCellValue("台中","B78")
// 底下處理
    xlsx.SetCellValue("台中","A71",strconv.Itoa(month)+"月目標業績："+sales_set_all)
    xlsx.SetCellValue("台中","B72",strconv.Itoa(month)+"月目標")


    //處理年度財務報表
for i:=3;i<=5;i++{
    for j:=1;j<13;j++{
        value:=now.GetCellValue("年度業績圖表",lib.Excel_int_to_label(j)+strconv.Itoa(i))
        var input interface{}
        intt,err:=strconv.Atoi(value)

        if err!=nil{
            input=value
        }else {
            input=intt

        }
        xlsx.SetCellValue("台中",lib.Excel_int_to_label(j)+strconv.Itoa(i),input)
    }

}
//處理季財務報表
    for i:=11;i<=13;i++{
        for j:=1;j<5;j++{
            value:=now.GetCellValue("年度業績圖表",lib.Excel_int_to_label(j)+strconv.Itoa(i))
            var input interface{}
            intt,err:=strconv.Atoi(value)

            if err!=nil{
                input=value
            }else {
                input=intt

            }
            xlsx.SetCellValue("台中",lib.Excel_int_to_label(j)+strconv.Itoa(i),input)
        }

    }
    //處理日報表

    xlsx.SetCellValue("台中","H19",strconv.Itoa(month)+"/"+strconv.Itoa(day)+"日報表")






    now_d,_:=lib.Get_day_info("./excel/"+now_filename,month,day)

    pass_d,_:=lib.Get_day_info("./excel/"+pass_filename,month,day)



    var d_gap_salse int
    var d_gap_commodity int
    var d_gap_tech int
    var d_gap_gift int
    var d_gap_salon int
    var d_gap_serve int

    //業績
    //當日業績設定
    t1,_:=strconv.Atoi(sales_set_all)
    d_gap_salse=t1/work_day
    xlsx.SetCellValue("台中","I21",d_gap_salse)
    //當日達成業績
    n_s,_:=strconv.Atoi(now_d.Sales)
    xlsx.SetCellValue("台中","I22",n_s)
    //去年當日業績
    p_s,_:=strconv.Atoi(pass_d.Sales)
    xlsx.SetCellValue("台中","I23",p_s)


    //商品
    t1,_=strconv.Atoi(commodity__set_all)
    d_gap_commodity=t1/work_day
    xlsx.SetCellValue("台中","I29",d_gap_commodity)
    //當日達成業績
    n_s,_=strconv.Atoi(now_d.Commodity)
    xlsx.SetCellValue("台中","I30",n_s)
    //去年當日業績
    p_s,_=strconv.Atoi(pass_d.Commodity)
    xlsx.SetCellValue("台中","I31",p_s)


    //技術回收
    t1,_=strconv.Atoi(tech_set_all)
    d_gap_tech=t1/work_day
    xlsx.SetCellValue("台中","I37",d_gap_tech)
    //當日達成業績
    n_s,_=strconv.Atoi(now_d.Tech)
    xlsx.SetCellValue("台中","I38",n_s)
    //去年當日業績
    p_s,_=strconv.Atoi(pass_d.Tech)
    xlsx.SetCellValue("台中","I39",p_s)


    //禮券
    t1,_=strconv.Atoi(gift_set_all)
    d_gap_gift=t1/work_day
    xlsx.SetCellValue("台中","I45",d_gap_gift)
    //當日達成業績
    n_s,_=strconv.Atoi(now_d.Gift)
    xlsx.SetCellValue("台中","I46",n_s)
    //去年當日業績
    p_s,_=strconv.Atoi(pass_d.Gift)
    xlsx.SetCellValue("台中","I47",p_s)

    //沙龍
    t1,_=strconv.Atoi(salon_set_all)
    d_gap_salon=t1/work_day
    xlsx.SetCellValue("台中","I53",d_gap_salon)
    //當日達成業績
    n_s,_=strconv.Atoi(now_d.Salon)
    xlsx.SetCellValue("台中","I54",n_s)
    //去年當日業績
    p_s,_=strconv.Atoi(pass_d.Salon)
    xlsx.SetCellValue("台中","I55",p_s)

    //服務單pay

    t1,_=strconv.Atoi(serve_set_all)
    d_gap_serve=t1/work_day
    xlsx.SetCellValue("台中","I61",d_gap_serve)
    //當日達成業績
    n_s,_=strconv.Atoi(now_d.Serve)
    xlsx.SetCellValue("台中","I62",n_s)
    //去年當日業績
    p_s,_=strconv.Atoi(pass_d.Serve)
    xlsx.SetCellValue("台中","I63",p_s)

//月報表
monin:=19
for i:=0;i<6;i++{
    xlsx.SetCellValue("台中","A"+strconv.Itoa(monin),strconv.Itoa(month)+"月週報")
    monin=monin+8
//日期
//
}

index:=21
for i:=0;i<6;i++{
    xlsx.SetCellValue("台中","A"+strconv.Itoa(index),"W1("+strconv.Itoa(month)+"/1~"+strconv.Itoa(month)+"/7)")
    xlsx.SetCellValue("台中","A"+strconv.Itoa(index+1),"W2("+strconv.Itoa(month)+"/8~"+strconv.Itoa(month)+"/14)")
    xlsx.SetCellValue("台中","A"+strconv.Itoa(index+2),"W3("+strconv.Itoa(month)+"/15~"+strconv.Itoa(month)+"/21)")
    xlsx.SetCellValue("台中","A"+strconv.Itoa(index+3),"W4("+strconv.Itoa(month)+"/22~"+strconv.Itoa(month)+"/28)")

index=index+8
    }



    month_day:=lib.Day_of_month(year,month)
index=25
    if month_day>28{

for i:=0;i<6;i++ {
    xlsx.SetCellValue("台中", "A"+strconv.Itoa(index), "W5("+strconv.Itoa(month)+"/29~"+strconv.Itoa(month)+"/"+strconv.Itoa(month_day)+")")




    style:=xlsx.GetCellStyle("台中","B"+strconv.Itoa(index-1))
    xlsx.SetCellStyle("台中","B"+strconv.Itoa(index),"B"+strconv.Itoa(index),style)

    style=xlsx.GetCellStyle("台中","C"+strconv.Itoa(index-1))
    xlsx.SetCellStyle("台中","C"+strconv.Itoa(index),"C"+strconv.Itoa(index),style)

    style=xlsx.GetCellStyle("台中","D"+strconv.Itoa(index-1))
    xlsx.SetCellStyle("台中","D"+strconv.Itoa(index),"D"+strconv.Itoa(index),style)


    xlsx.SetCellFormula("台中","D"+strconv.Itoa(index),"=SUM(B"+strconv.Itoa(index)+"/C"+strconv.Itoa(index)+")")

    style=xlsx.GetCellStyle("台中","E"+strconv.Itoa(index-1))
    xlsx.SetCellStyle("台中","E"+strconv.Itoa(index),"E"+strconv.Itoa(index),style)

    style=xlsx.GetCellStyle("台中","F"+strconv.Itoa(index-1))
    xlsx.SetCellStyle("台中","F"+strconv.Itoa(index),"F"+strconv.Itoa(index),style)
    xlsx.SetCellFormula("台中","F"+strconv.Itoa(index),"=SUM(B"+strconv.Itoa(index)+"/E"+strconv.Itoa(index)+")")

    style=xlsx.GetCellStyle("台中","G"+strconv.Itoa(index-1))
    xlsx.SetCellStyle("台中","G"+strconv.Itoa(index),"G"+strconv.Itoa(index),style)



    index = index + 8
}
    }



//月報目標設定



shift:=21



    for i:=1;i<=4;i++{
        w_i:=lib.Get_week_info("./excel/"+now_filename,month,i)
    //今年業績
        w_sales,_:= strconv.Atoi(w_i.Sales)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift),w_sales)

        w_com,_:= strconv.Atoi(w_i.Commodity)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+8),w_com)

        w_tech,_:= strconv.Atoi(w_i.Tech)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+16),w_tech)

        w_gift,_:= strconv.Atoi(w_i.Gift)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+24),w_gift)

        w_salon,_:= strconv.Atoi(w_i.Salon)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+32),w_salon)

        w_serve,_:= strconv.Atoi(w_i.Serve)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+40),w_serve)

        //去年業績

        w_i=lib.Get_week_info("./excel/"+pass_filename,month,i)

        w_sales,_= strconv.Atoi(w_i.Sales)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift),w_sales)

        w_com,_= strconv.Atoi(w_i.Commodity)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+8),w_com)

        w_tech,_= strconv.Atoi(w_i.Tech)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+16),w_tech)

        w_gift,_= strconv.Atoi(w_i.Gift)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+24),w_gift)

        w_salon,_= strconv.Atoi(w_i.Salon)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+32),w_salon)

        w_serve,_= strconv.Atoi(w_i.Serve)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+40),w_serve)

        //目標設定


        w_ss:= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_salse
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_commodity
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+8),w_ss)


        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_tech
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+16),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_gift
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+24),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_salon
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+32),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_serve
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+40),w_ss)





        shift=shift+1
    }



i:=5
    if month_day>28{
        w_i:=lib.Get_week_info("./excel/"+now_filename,month,i)
        //今年業績
        w_sales,_:= strconv.Atoi(w_i.Sales)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift),w_sales)

        w_com,_:= strconv.Atoi(w_i.Commodity)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+8),w_com)

        w_tech,_:= strconv.Atoi(w_i.Tech)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+16),w_tech)

        w_gift,_:= strconv.Atoi(w_i.Gift)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+24),w_gift)

        w_salon,_:= strconv.Atoi(w_i.Salon)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+32),w_salon)

        w_serve,_:= strconv.Atoi(w_i.Serve)
        xlsx.SetCellValue("台中","B"+strconv.Itoa(shift+40),w_serve)

        //去年業績

        w_i=lib.Get_week_info("./excel/"+pass_filename,month,i)

        w_sales,_= strconv.Atoi(w_i.Sales)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift),w_sales)

        w_com,_= strconv.Atoi(w_i.Commodity)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+8),w_com)

        w_tech,_= strconv.Atoi(w_i.Tech)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+16),w_tech)

        w_gift,_= strconv.Atoi(w_i.Gift)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+24),w_gift)

        w_salon,_= strconv.Atoi(w_i.Salon)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+32),w_salon)

        w_serve,_= strconv.Atoi(w_i.Serve)
        xlsx.SetCellValue("台中","E"+strconv.Itoa(shift+40),w_serve)

        //目標設定


        w_ss:= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_salse
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_commodity
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+8),w_ss)


        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_tech
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+16),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_gift
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+24),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_salon
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+32),w_ss)

        w_ss= lib.Day_of_week_work("./excel/"+now_filename,month,i)*d_gap_serve
        xlsx.SetCellValue("台中","C"+strconv.Itoa(shift+40),w_ss)



    }









    //刪除控制碼
    for i:=82;i<=86;i++{
        for j:=1;j<7;j++{
          //  fmt.Println(excel_col[j],lib.Excel_int_to_label(j),j)
            xlsx.SetCellValue("台中",lib.Excel_int_to_label(j)+strconv.Itoa(i),"")
            xlsx.SetCellStyle("台中",lib.Excel_int_to_label(j)+strconv.Itoa(i),lib.Excel_int_to_label(j)+strconv.Itoa(i),0)
        }

    }
xlsx.SetSheetName("台中","台中"+strconv.Itoa(year))
xlsx.SaveAs("./完成輸出檔案/"+strconv.Itoa(year)+"年度達成進度.xlsx")



fmt.Println("轉換完成")
}