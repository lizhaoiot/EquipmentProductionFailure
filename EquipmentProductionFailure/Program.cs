using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SQL;
using System.Data;
using System.Data.SqlClient;
using System.Threading;

namespace EquipmentProductionFailure
{
    class Program
    {
        private static string CurrentYearQJ = string.Empty;
        private static string CurrentMonthQJ = string.Empty;
        private static string CurrentDayQJ = string.Empty;
        private static string CurrentMonthFirstQJ = string.Empty;
        private static string CurrentMonthlLastQJ = string.Empty;
        private static double GetSecondsQJ = 0;
        static void Main(string[] args)
        {
            String CurrentYears = DateTime.Now.Year.ToString();
            String CurrentMonths = DateTime.Now.Month.ToString();
            DateTime CurrentDays =Convert.ToDateTime( "2021-1-31"); //
            //DateTime CurrentDays = DateTime.Now;
            //每月最后一天不更新
            for (int i = 0; i <Convert.ToInt32(CurrentDays.Day.ToString()); i++)
            {
                SqlParameter[] para ={
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
              };
                para[0].Value = CurrentYears;
                para[1].Value= CurrentDays.AddMonths(-1).Month.ToString();
                DataTable DtTime = SqlHelper.ExecStoredProcedureDataTable("GetMonthsFirstLastDay", para);
                String CurrentMonthFirst =Convert.ToDateTime(DtTime.Rows[0][1].ToString()).AddDays(-1).ToString("d") +" 8:30:00";
                String CurrentMonthlLast = CurrentDays.AddDays(-i).ToString("d") + " 8:30:00";
                //TimeSpan t = Convert.ToDateTime(CurrentMonthlLast) - Convert.ToDateTime(CurrentMonthFirst);
                TimeSpan t = Convert.ToDateTime(CurrentMonthlLast) - Convert.ToDateTime("2020/12/31 8:30:00");
                
                double GetSeconds = t.TotalSeconds;//获取时间差相差的秒数

                SqlHelper.ExecCommand("DELETE FROM MonthlyEquipmentFailure WHERE Years='" + CurrentYears + "' AND Months='" + CurrentMonths + "' AND Days='" +Convert.ToDateTime( CurrentMonthlLast).Day.ToString() + "'");

                CurrentYearQJ = CurrentYears;
                CurrentMonthQJ = CurrentMonths;
                CurrentDayQJ = Convert.ToDateTime(CurrentMonthlLast).Day.ToString();
                //CurrentMonthFirstQJ = CurrentMonthFirst;
                CurrentMonthFirstQJ = "2020/12/31 8:30:00";
                
                CurrentMonthlLastQJ = CurrentMonthlLast;
                GetSecondsQJ = GetSeconds;

                //杭州两片罐
                try
                {
                   ExecSQLHZ(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//杭州
                }
                catch (Exception ex)
                {
                    Console.WriteLine("计算杭州两片罐数据发生错误" + ex.Message);
                }

                //天津两片罐
                try
                {
                    ExecSQLTJ(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//天津
                }
                catch (Exception ex)
                {
                    Console.WriteLine("计算天津两片罐数据发生错误" + ex.Message);
                }

                //纪鸿两片罐
                try
                { 
                   ExecSQLJH(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//纪鸿
                }
                catch (Exception ex)
                {
                    Console.WriteLine("计算纪鸿两片罐数据发生错误" + ex.Message);
                }

                //武汉两片罐
                try
                {

                    ExecSQLWH(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//武汉
                }
                catch (Exception ex)
                {

                    Console.WriteLine("计算武汉两片罐数据发生错误" + ex.Message);
                }

                //广州两片罐
                try
                {
                    ExecSQLGZ(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//广州
                }
                catch (Exception ex )
                {

                    Console.WriteLine("计算广州两片罐数据发生错误" + ex.Message);
                }

                //南宁两片罐
                try
                {
                    ExecSQLNN(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//南宁
                }
                catch (Exception ex)
                {

                    Console.WriteLine("计算南宁两片罐数据发生错误" + ex.Message);
                }

                //龙泉两片罐
                try
                {
                   ExecSQLLQ(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//龙泉
                }//
                catch (Exception ex)
                {

                    Console.WriteLine("计算龙泉两片罐数据发生错误" + ex.Message);

                }

                //温江两片罐
                try
                {
                    ExecSQLWJ(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//温江
                }
                catch (Exception ex)
                {

                    Console.WriteLine("计算温江两片罐数据发生错误" + ex.Message);
                }

                //福建两片罐

                try
                {//
                   ExecSQLFJ(CurrentYearQJ, CurrentMonthQJ, CurrentDayQJ, CurrentMonthFirstQJ, CurrentMonthlLastQJ, GetSecondsQJ);//福建
                }
                catch (Exception ex)
                {

                    Console.WriteLine("计算福建两片罐数据发生错误" + ex.Message);
                }



            }

        }



        /// <summary>
        /// 杭州
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLHZ(string CurrentYear, string CurrentMonth, string CurrentDay,string CurrentMonthFirst,string CurrentMonthlLast,double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure (FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('7000','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','" + CurrentDay + "')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "7000";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "7000";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth +  "'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth +  "'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "7000";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "7000";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "7000";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "7000";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "7000";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "7000";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='7000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "7000";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "7000";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "7000";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "7000";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("7000", CurrentYear, CurrentMonth,Convert.ToDateTime( CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 天津
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLTJ(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {

            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('5000','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "5000";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "5000";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "5000";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "5000";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "5000";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "5000";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "5000";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "5000";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='5000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "5000";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "5000";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "5000";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "5000";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("5000", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }

        /// <summary>
        /// 纪鸿
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLJH(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('HN00','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "HN00";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "HN00";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "HN00";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "HN00";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "HN00";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "HN00";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "HN00";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "HN00";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='HN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "HN00";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "HN00";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "HN00";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "HN00";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("HN00", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 武汉
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLWH(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('9000','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "9000";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "9000";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "9000";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "9000";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "9000";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "9000";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "9000";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "9000";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='9000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "9000";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "9000";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "9000";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "9000";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("9000", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 广州
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLGZ(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('2010','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "2010";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "2010";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "2010";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "2010";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "2010";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "2010";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "2010";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "2010";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='2010' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "2010";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "2010";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "2010";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "2010";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("2010", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 南宁
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLNN(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('NN00','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "NN00";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "NN00";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //底涂
            SqlParameter[] para22 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "NN00";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "H";
            DataTable dtPaddingMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET PaddingMachineX='" + dtPaddingMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET PaddingMachineX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtPaddingMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET PaddingMachineY='" + dtPaddingMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET PaddingMachineY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtPaddingMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET PaddingMachineZ='" + dtPaddingMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET PaddingMachineZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "NN00";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "NN00";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "NN00";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "NN00";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //彩印
            SqlParameter[] para61 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para61[0].Value = "NN00";
            para61[1].Value = CurrentMonthFirst;
            para61[2].Value = CurrentMonthlLast;
            para61[3].Value = "Z";

            DataTable ColorPrintMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para61);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineX='" + ColorPrintMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ColorPrintMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para61);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineY='" + ColorPrintMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ColorPrintMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para61);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineZ='" + ColorPrintMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "NN00";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "NN00";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='NN00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "NN00";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "NN00";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "NN00";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            UpdateWorkHoursActual("NN00", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 龙泉
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLLQ(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('4020','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "4020";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "4020";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "4020";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "4020";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "4020";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "4020";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //彩印
            SqlParameter[] para61 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para61[0].Value = "4020";
            para61[1].Value = CurrentMonthFirst;
            para61[2].Value = CurrentMonthlLast;
            para61[3].Value = "Z";

            DataTable ColorPrintMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para61);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineX='" + ColorPrintMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ColorPrintMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para61);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineY='" + ColorPrintMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ColorPrintMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para61);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineZ='" + ColorPrintMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ColorPrintMachineZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "4020";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "4020";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='4020' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "4020";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "4020";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "4020";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);

            UpdateWorkHoursActual("4020", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 温江
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLWJ(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('4000','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "4000";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "4000";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "4000";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "4000";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "4000";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "4000";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "4000";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "4000";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='4000' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "4000";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "4000";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "4000";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "4000";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("4000", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        /// <summary>
        /// 福建
        /// </summary>
        /// <param name="CurrentYear"></param>
        /// <param name="CurrentMonth"></param>
        /// <param name="CurrentDay"></param>
        private static void ExecSQLFJ(string CurrentYear, string CurrentMonth, string CurrentDay, string CurrentMonthFirst, string CurrentMonthlLast, double GetSeconds)
        {
            //如果不存在就插入\
            SqlHelper.ExecCommand("INSERT INTO MonthlyEquipmentFailure(FactoryID,Years,Months,WorkHoursTheory,WorkHoursActual,CupMachineX,CupMachineY,CupMachineZ,DrawingMachineX,DrawingMachineY,DrawingMachineZ,WashCansMachineX,WashCansMachineY,WashCansMachineZ,PaddingMachineX,PaddingMachineY,PaddingMachineZ,ColorPrintMachineX,ColorPrintMachineY,ColorPrintMachineZ,CoaterMachineX,CoaterMachineY,CoaterMachineZ,IBOX,IBOY,IBOZ,ShrinkageTurningX,ShrinkageTurningY,ShrinkageTurningZ,StackerCraneX,StackerCraneY,StackerCraneZ,DriveWireX,DriveWireY,DriveWireZ,AuxiliaryX,AuxiliaryY,AuxiliaryZ,Days) VALUES('FJ00','" + CurrentYear + "','" + CurrentMonth + "','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','"+CurrentDay+"')");
            SqlParameter[] para1 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar)
              };
            para1[0].Value = "FJ00";
            para1[1].Value = CurrentMonthFirst;
            para1[2].Value = CurrentMonthlLast;
            DataTable dtWorkHoursTheory = SqlHelper.ExecStoredProcedureDataTable("CalcTheoryWorkingHours", para1);
            //理论工时
            if (dtWorkHoursTheory.Rows.Count > 0)
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull(dtWorkHoursTheory.Rows[0][0].ToString()), GetSeconds.ToString()) + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            else
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WorkHoursTheory='" + CalcSpanhours(JudgeISNull("0"), GetSeconds.ToString()) + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //计算单机台故障停机时间        
            //冲杯
            SqlParameter[] para2 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para2[0].Value = "FJ00";
            para2[1].Value = CurrentMonthFirst;
            para2[2].Value = CurrentMonthlLast;
            para2[3].Value = "A";
            DataTable dtCupMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='" + dtCupMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='" + dtCupMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable dtCupMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para2);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='" + dtCupMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET CupMachineZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //清洗
            SqlParameter[] para3 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
              };
            para3[0].Value = "FJ00";
            para3[1].Value = CurrentMonthFirst;
            para3[2].Value = CurrentMonthlLast;
            para3[3].Value = "B";
            DataTable WashCansMachineX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='" + WashCansMachineX.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='" + WashCansMachineY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable WashCansMachineZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para3);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='" + WashCansMachineZ.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET WashCansMachineZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //IBO
            SqlParameter[] para4 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para4[0].Value = "FJ00";
            para4[1].Value = CurrentMonthFirst;
            para4[2].Value = CurrentMonthlLast;
            para4[3].Value = "C";

            DataTable IBOX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='" + IBOX.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='" + IBOY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable IBOZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para4);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='" + IBOZ.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET IBOZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }


            //缩翻
            SqlParameter[] para5 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para5[0].Value = "FJ00";
            para5[1].Value = CurrentMonthFirst;
            para5[2].Value = CurrentMonthlLast;
            para5[3].Value = "D";

            DataTable ShrinkageTurningX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='" + ShrinkageTurningX.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='" + ShrinkageTurningY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable ShrinkageTurningZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para5);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='" + ShrinkageTurningZ.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET ShrinkageTurningZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }

            //码垛
            SqlParameter[] para6 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para6[0].Value = "FJ00";
            para6[1].Value = CurrentMonthFirst;
            para6[2].Value = CurrentMonthlLast;
            para6[3].Value = "E";

            DataTable StackerCraneX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='" + StackerCraneX.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='" + StackerCraneY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable StackerCraneZ = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeZ", para6);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='" + StackerCraneZ.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET StackerCraneZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //线控
            SqlParameter[] para7 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para7[0].Value = "FJ00";
            para7[1].Value = CurrentMonthFirst;
            para7[2].Value = CurrentMonthlLast;
            para7[3].Value = "F";
            DataTable DriveWireY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para7);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='" + DriveWireY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET DriveWireZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //辅机
            SqlParameter[] para8 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para8[0].Value = "FJ00";
            para8[1].Value = CurrentMonthFirst;
            para8[2].Value = CurrentMonthlLast;
            para8[3].Value = "G";
            SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryZ='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            DataTable AuxiliaryX = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeX", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='" + AuxiliaryX.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryX='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            DataTable AuxiliaryY = SqlHelper.ExecStoredProcedureDataTable("CalcSingleStationDowntimeY", para8);
            try
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='" + AuxiliaryY.Rows[0][0].ToString() + "' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            catch
            {
                SqlHelper.ExecCommand("UPdate MonthlyEquipmentFailure SET AuxiliaryY='0' WHERE FactoryID='FJ00' AND Years='" + CurrentYear + "' AND Months='" + CurrentMonth + "'"+" AND Days='"+CurrentDay+"'");
            }
            //计算多机台故障停机时间
            //拉伸
            SqlParameter[] para9 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para9[0].Value = "FJ00";
            para9[1].Value = CurrentYear;
            para9[2].Value = CurrentMonth;
            para9[3].Value = CurrentMonthFirst;
            para9[4].Value = CurrentMonthlLast;
            para9[5].Value = "A";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para9);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para9);

            //底涂
            SqlParameter[] para10 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para10[0].Value = "FJ00";
            para10[1].Value = CurrentYear;
            para10[2].Value = CurrentMonth;
            para10[3].Value = CurrentMonthFirst;
            para10[4].Value = CurrentMonthlLast;
            para10[5].Value = "B";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para10);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para10);
            //内涂
            SqlParameter[] para11 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para11[0].Value = "FJ00";
            para11[1].Value = CurrentYear;
            para11[2].Value = CurrentMonth;
            para11[3].Value = CurrentMonthFirst;
            para11[4].Value = CurrentMonthlLast;
            para11[5].Value = "C";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para11);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para11);
            //彩印
            SqlParameter[] para12 ={
                                     new SqlParameter("@FactoryID",SqlDbType.VarChar),
                                     new SqlParameter("@Years",SqlDbType.VarChar),
                                     new SqlParameter("@Months",SqlDbType.VarChar),
                                     new SqlParameter("@FirstDay",SqlDbType.VarChar),
                                     new SqlParameter("@LastDay",SqlDbType.VarChar),
                                     new SqlParameter("@Flag",SqlDbType.VarChar),
               };
            para12[0].Value = "FJ00";
            para12[1].Value = CurrentYear;
            para12[2].Value = CurrentMonth;
            para12[3].Value = CurrentMonthFirst;
            para12[4].Value = CurrentMonthlLast;
            para12[5].Value = "D";
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureX", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureY", para12);
            SqlHelper.ExecStoredProcedure("CalcMultiMachineEquipmentFailureZ", para12);
            UpdateWorkHoursActual("FJ00", CurrentYear, CurrentMonth, Convert.ToDateTime(CurrentMonthlLast).Day.ToString());
        }
        //两个时间相差多少个小时
        private static String CalcSpanhours(String s1,String s2)
        {
            double d1 = Convert.ToDouble(s1);
            double d2 = Convert.ToDouble(s2);
            double s = (d2 - d1)/3600;
            return s.ToString();
        }
        private static string JudgeISNull(String s1)
        {
            if (String.IsNullOrEmpty(s1))
            {
                return "0";
            }
            else
            {
                return s1;
            }
        }
        private int DateDiff(DateTime dateStart, DateTime dateEnd)
        {
            DateTime start = Convert.ToDateTime(dateStart.ToShortDateString());
            DateTime end = Convert.ToDateTime(dateEnd.ToShortDateString());
            TimeSpan sp = end.Subtract(start);
            return sp.Days;
        }
        private static void UpdateWorkHoursActual(String FactoryID , String Years , String Months, String Days)
        {
            string sqlstr = "SELECT * FROM MonthlyEquipmentFailure WHERE FactoryID='" + FactoryID + "' AND Years='" + Years + "' AND Months='" + Months + "' AND Days='" + Days + "'";
            DataTable T1 = SqlHelper.ExecuteDataTable(sqlstr);
            String WorkHoursTheory = T1.Rows[0]["WorkHoursTheory"].ToString();
            String CupMachineX = T1.Rows[0]["CupMachineX"].ToString();
            String CupMachineY = T1.Rows[0]["CupMachineY"].ToString();
            String CupMachineZ = T1.Rows[0]["CupMachineZ"].ToString();
            String DrawingMachineX = T1.Rows[0]["DrawingMachineX"].ToString();
            String DrawingMachineY = T1.Rows[0]["DrawingMachineY"].ToString();
            String DrawingMachineZ = T1.Rows[0]["DrawingMachineZ"].ToString();
            String WashCansMachineX = T1.Rows[0]["WashCansMachineX"].ToString();
            String WashCansMachineY = T1.Rows[0]["WashCansMachineY"].ToString();
            String WashCansMachineZ = T1.Rows[0]["WashCansMachineZ"].ToString();
            String PaddingMachineX = T1.Rows[0]["PaddingMachineX"].ToString();
            String PaddingMachineY = T1.Rows[0]["PaddingMachineY"].ToString();
            String PaddingMachineZ = T1.Rows[0]["PaddingMachineZ"].ToString();
            String ColorPrintMachineX = T1.Rows[0]["ColorPrintMachineX"].ToString();
            String ColorPrintMachineY = T1.Rows[0]["ColorPrintMachineY"].ToString();
            String ColorPrintMachineZ = T1.Rows[0]["ColorPrintMachineZ"].ToString();
            String CoaterMachineX = T1.Rows[0]["CoaterMachineX"].ToString();
            String CoaterMachineY = T1.Rows[0]["CoaterMachineY"].ToString();
            String CoaterMachineZ = T1.Rows[0]["CoaterMachineZ"].ToString();
            String IBOX = T1.Rows[0]["IBOX"].ToString();
            String IBOY = T1.Rows[0]["IBOY"].ToString();
            String IBOZ = T1.Rows[0]["IBOZ"].ToString();
            String ShrinkageTurningX = T1.Rows[0]["ShrinkageTurningX"].ToString();
            String ShrinkageTurningY = T1.Rows[0]["ShrinkageTurningY"].ToString();
            String ShrinkageTurningZ = T1.Rows[0]["ShrinkageTurningZ"].ToString();
            String StackerCraneX = T1.Rows[0]["StackerCraneX"].ToString();
            String StackerCraneY = T1.Rows[0]["StackerCraneY"].ToString();
            String StackerCraneZ = T1.Rows[0]["StackerCraneZ"].ToString();
            String DriveWireY = T1.Rows[0]["DriveWireY"].ToString();
            String AuxiliaryX = T1.Rows[0]["AuxiliaryX"].ToString();
            String AuxiliaryY = T1.Rows[0]["AuxiliaryY"].ToString();
            String WorkHoursActual = Convert.ToString((ConvertDouble(WorkHoursTheory)- ConvertDouble(CupMachineX) - ConvertDouble(CupMachineY) - ConvertDouble(CupMachineZ) - ConvertDouble(DrawingMachineX) - ConvertDouble(DrawingMachineY) - ConvertDouble(DrawingMachineZ) - ConvertDouble(WashCansMachineX) - ConvertDouble(WashCansMachineY) - ConvertDouble(WashCansMachineZ) - ConvertDouble(PaddingMachineX) - ConvertDouble(PaddingMachineY) - ConvertDouble(PaddingMachineZ) - ConvertDouble(ColorPrintMachineX) - ConvertDouble(ColorPrintMachineY) - ConvertDouble(ColorPrintMachineZ) - ConvertDouble(CoaterMachineX) - ConvertDouble(CoaterMachineY) - ConvertDouble(CoaterMachineZ) - ConvertDouble(IBOX) - ConvertDouble(IBOY) - ConvertDouble(IBOZ) - ConvertDouble(ShrinkageTurningX) - ConvertDouble(ShrinkageTurningY) - ConvertDouble(ShrinkageTurningZ) - ConvertDouble(StackerCraneX) - ConvertDouble(StackerCraneY) - ConvertDouble(StackerCraneZ) - ConvertDouble(DriveWireY) - ConvertDouble(AuxiliaryX) - ConvertDouble(AuxiliaryY)));
            SqlHelper.ExecCommand("UPDATE MonthlyEquipmentFailure SET WorkHoursActual='"+ WorkHoursActual + "' WHERE FactoryID='"+ FactoryID + "' AND Years='"+ Years + "' AND Months='"+ Months + "' AND Days='" + Days + "'");
        }
        private static Decimal ConvertDouble(String str)
        {
            try
            {
               if (String.IsNullOrEmpty(str))
               {
                   return 0;
               }
               else
               {
                 return ChangeToDecimal(str);
               }
            }
            catch (Exception)
            {
                return 0;
            }

        }
        /// <summary>
        /// 数字科学计数法处理
        /// </summary>
        /// <param name="strData"></param>
        /// <returns></returns>
        private static Decimal ChangeToDecimal(string strData)
        {
            Decimal dData = 0.0M;
            if (strData.Contains("E"))
            {
                dData = Convert.ToDecimal(Decimal.Parse(strData.ToString(), System.Globalization.NumberStyles.Float));
            }
            else
            {
                dData = Convert.ToDecimal(strData);
            }
            return Math.Round(dData, 4);
        }

    }
}
