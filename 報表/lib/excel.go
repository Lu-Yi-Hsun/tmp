package lib
import(
	"github.com/360EntSecGroup-Skylar/excelize"
	"fmt"
	"strconv"
	"time"


)

var excel_col=[...]string{"Z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y"}


func Get_week_info(filename string,month int,week int )Day_info{
	ans:=Day_info{}
	w:=week
	d:=1+(7*(w-1))


	week_day:=0
	for i:=d;i<d+7;i++ {

		info,err:=Get_day_info(filename, month, i)


		if err==nil{


			i, _ := strconv.Atoi(info.Salon)
			j, _ := strconv.Atoi(ans.Salon)
			ans.Salon=strconv.Itoa(j+i)

			i, _ = strconv.Atoi(info.Commodity)
			j, _ = strconv.Atoi(ans.Commodity)
			ans.Commodity=strconv.Itoa(j+i)



			i, _ = strconv.Atoi(info.Tech)
			j, _ = strconv.Atoi(ans.Tech)
			ans.Tech=strconv.Itoa(j+i)


			i, _ = strconv.Atoi(info.Serve)
			j, _ = strconv.Atoi(ans.Serve)
			ans.Serve=strconv.Itoa(j+i)


			i, _ = strconv.Atoi(info.Gift)
			j, _ = strconv.Atoi(ans.Gift)
			ans.Gift=strconv.Itoa(j+i)


			i, _ = strconv.Atoi(info.Sales)
			j, _ = strconv.Atoi(ans.Sales)
			ans.Sales=strconv.Itoa(j+i)







			week_day=week_day+1
		}
	}


	return ans
}


func Day_of_week_work(filename string,month int,week int)int{
	w:=week
	d:=1+(7*(w-1))


	week_day:=0
	for i:=d;i<d+7;i++ {

		_,err:=Get_day_info(filename, month, i)

		if err==nil{
			week_day=week_day+1
		}
	}
return week_day



}

func Day_of_month(year int,month int)int {
month=month+1
	switch month {

	case 1:
		return time.Date(year, 1, 0, 0, 0, 0, 0, time.UTC).Day()
	case 2:
		return time.Date(year, 2, 0, 0, 0, 0, 0, time.UTC).Day()
	case 3:
		return time.Date(year, 3, 0, 0, 0, 0, 0, time.UTC).Day()
	case 4:
		return time.Date(year, 4, 0, 0, 0, 0, 0, time.UTC).Day()
	case 5:
		return time.Date(year, 5, 0, 0, 0, 0, 0, time.UTC).Day()
	case 6:
		return time.Date(year, 6, 0, 0, 0, 0, 0, time.UTC).Day()
	case 7:
		return time.Date(year, 7, 0, 0, 0, 0, 0, time.UTC).Day()
	case 8:
		return time.Date(year, 8, 0, 0, 0, 0, 0, time.UTC).Day()
	case 9:
		return time.Date(year, 9, 0, 0, 0, 0, 0, time.UTC).Day()
	case 10:
		return time.Date(year, 10, 0, 0, 0, 0, 0, time.UTC).Day()
	case 11:
		return time.Date(year, 11, 0, 0, 0, 0, 0, time.UTC).Day()
	case 12:
		return time.Date(year, 12, 0, 0, 0, 0, 0, time.UTC).Day()
	}


return -1




}
func Excel_int_to_label(number int)string{

	ans:=""
	if number==0{

		return excel_col[1]
	}else if(number%26==0){

		for {
			if number <= 0 {
				break
			}
			if number%26==0 {
				ans = excel_col[0] + ans
			}else {
				ans = excel_col[number%26-1] + ans
			}

			number = number / 26


		}

		return ans

	}else {

		for {
			if number <= 0 {
				break
			}

			ans =  excel_col[number%26]+ans

			number = number / 26


		}

		return ans

	}
}



type Day_info struct {
	//業績
	Sales string
	//商品
	Commodity string
	//技術回收
	Tech string
	//禮券
	Gift string
	//沙龍
	Salon string
	//服務單pay
	Serve string

}



type day_shift struct {
	//業績
	sales int
	//商品
	commodity int
	//技術回收
	tech int
	//禮券
	gift int
	//沙龍
	salon int
	//服務單pay
	serve int

}

func excel_time_to_unix(Excel_Timestamp int64)int64{
	return (Excel_Timestamp - 25569) * 86400

}

func Get_month_day(cell string)(int,int,int){


	i, _ := strconv.ParseInt(cell, 10, 64);
	tm := time.Unix(excel_time_to_unix(i),0)


	return tm.Year(),int(tm.Month()),tm.Day()
}
type error interface {
	Error() string
}

func Get_day_info(file string,month int,day int)(Day_info,error){
	sheet:="營業報"
	sheet_tech:="技術回收"
	xlsx, err := excelize.OpenFile(file)
	ans:=Day_info{Sales:"0",Commodity:"0",Tech:"0",Gift:"0",Salon:"0",Serve:"0"}

	if err != nil {


		fmt.Println(err)
		return ans,fmt.Errorf("無法開啟檔案")
	}
	// Get value from cell by given worksheet name and axis.

//保護區塊 目的盡量保護	excel的破獲


var right_day day_shift
//這邊來控制強至找尋100位置
control_find:=70
for protect:=1;protect<control_find;protect++{
	cell := xlsx.GetCellValue(sheet, Excel_int_to_label(protect)+"3")

		switch cell {


		case "商品零售":
			right_day.commodity=protect

		case "禮券銷售":
			right_day.gift=protect
		case "每日業績":
			right_day.sales=protect
		case "療程服務":
			right_day.serve=protect


		}

}
	for protect:=1;protect<control_find;protect++{
		cell := xlsx.GetCellValue(sheet, Excel_int_to_label(protect)+"4")

		switch cell {

		case "沙2":
			right_day.salon=protect


		}

	}




	for protect:=1;protect<control_find;protect++{
		cell := xlsx.GetCellValue(sheet_tech, Excel_int_to_label(protect)+"3")





			if cell=="營業報"{
				right_day.tech=protect}



	}



//
index:=5
for {
	cell := xlsx.GetCellValue(sheet, "A"+strconv.Itoa(index))

	if cell==""{

		return ans,fmt.Errorf("找不到該時間")
		break
	}
	_,m,d:=Get_month_day(cell)

	if m==month&&d==day{
		//fmt.Println("成功找到時間")
		//業績
		if  xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.sales)+strconv.Itoa(index))!="" {
			ans.Sales = xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.sales)+strconv.Itoa(index))
		}
		//商品

		M, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.commodity-2)+strconv.Itoa(index)))
		N, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.commodity-3)+strconv.Itoa(index)))
		ans.Commodity=strconv.Itoa(M+N)

		//技術回收
		if xlsx.GetCellValue(sheet_tech, Excel_int_to_label(right_day.tech)+strconv.Itoa(index))!="" {
			ans.Tech = xlsx.GetCellValue(sheet_tech, Excel_int_to_label(right_day.tech)+strconv.Itoa(index))
		}
		//禮券
		if xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.gift)+strconv.Itoa(index))!=""{
			ans.Gift = xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.gift)+strconv.Itoa(index))
		}

		//沙龍
		E, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.salon)+strconv.Itoa(index)))
		F, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.salon+1)+strconv.Itoa(index)))
		G, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.salon+2)+strconv.Itoa(index)))
		I, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.salon+4)+strconv.Itoa(index)))
		J, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.salon+5)+strconv.Itoa(index)))
		K, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.salon+6)+strconv.Itoa(index)))
		ans.Salon=strconv.Itoa(E+F+G+I+J+K)

		//服務單pay
		Q, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.serve-2)+strconv.Itoa(index)))
		R, _ := strconv.Atoi(xlsx.GetCellValue(sheet, Excel_int_to_label(right_day.serve-3)+strconv.Itoa(index)))
		ans.Serve=strconv.Itoa(Q+R)
		break





	}
index=index+1

}

	return ans,nil
	// Get all the rows in the Sheet1.


}

func Get_work_day_for_week(file string,){



}




func Get_work_day(file string)(int,error){
	//拿到這個月幾天
	ans:=0
	sheet:="營業報"

	xlsx, err := excelize.OpenFile(file)


	if err != nil {


		fmt.Println(err)
		return ans,fmt.Errorf("無法開啟檔案")
	}


	index:=5
	for {
		cell := xlsx.GetCellValue(sheet, "A"+strconv.Itoa(index))

		if cell==""{


			break
		}
		ans=ans+1
		index=index+1

	}

return ans,nil

}