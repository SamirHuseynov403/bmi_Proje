using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMI.Classlar.Emek_Haqqi
{
    internal class cl_EmekHaqqiHesablama
    {
        decimal[] salaries = { 400, 440, 480 }; // Maaş artımı dövrləri
        int[] months = { 4, 4, 4 }; // Hər maaş dövründə işlənmiş ayların sayı
        int vacationDays = 56; // Məzuniyyət günlərinin sayı
        
        public static decimal VacationPay(decimal SalaryLast12Months, int months, int vacationday)
        {
            decimal OAAH = Math.Round(SalaryLast12Months / months, 2);
            decimal BGAH = Math.Round(OAAH / 30.4m, 2);
            decimal MH = Math.Round(vacationday * BGAH, 2);
            return MH;
        }
        public static decimal CalculateVacationPayWithFlexibleIndexing(decimal[] salaries, int[] months, int vacationDays)
        {
            if (salaries.Length != months.Length)
                throw new ArgumentException("Maaş və ay sayı eyni olmalıdır!");

            decimal totalSalary = 0;
            int totalMonths = 12;

            for (int i = 0; i < salaries.Length; i++)
            {
                decimal indexedSalary = salaries[i];

                if (i > 0)
                {
                    decimal indexFactor = salaries[i] / salaries[i - 1];

                    if (indexFactor > 1)
                    {
                        indexedSalary = salaries[i - 1] * indexFactor;
                    }
                    else
                    {
                        indexedSalary = salaries[i];
                    }
                }

                totalSalary += indexedSalary * months[i];
            }

            decimal averageMonthlySalary = totalSalary / totalMonths;
            decimal dailySalary = averageMonthlySalary / 30.4m;
            return dailySalary * vacationDays;
        }
        public static decimal SickPay(int staj, decimal SalaryLast12Months, int Last12Monthsdays, int sickkday)
        {
            decimal staj_8 = 0.6m;
            decimal staj_12 = 0.8m;
            decimal staj_12Plus = 1m;
            decimal sickpay = 0;
            if (staj < 8)
            {
                sickpay = (SalaryLast12Months / Last12Monthsdays * sickkday) * staj_8;
            }

            else if (staj >= 8 && staj < 12)
            {
                sickpay = (SalaryLast12Months / Last12Monthsdays * sickkday) * staj_12;
            }

            else if (staj >= 12)
            {
                sickpay = (SalaryLast12Months / Last12Monthsdays * sickkday) * staj_12Plus;
            }

            return Math.Round(sickpay, 2);
        }
    }
}
