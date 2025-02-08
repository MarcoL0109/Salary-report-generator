import datetime
import calendar



class Calculator:
    def __init__(self):
        self.basic_salary = None
        self.join_date = None
        self.sick_leave_taken = None
        self.no_paid_leave_taken = None
        self.resigning_date = None
        self.resigning = None
        self.annual_leave = None
        self.paid_date = None
        self.curr_date = datetime.date.today()
        first = self.curr_date.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        self.last_month = datetime.date(last_month.year, last_month.month, 1)


    def update_values(self, basic_salary=None, join_date=None, sick_leave_taken=None, no_paid_leave_taken=None, resigning_date=None, resigning=None, annual_leave=None, paid_date=None):
        if basic_salary is not None:
            self.basic_salary = basic_salary
        if sick_leave_taken is not None:
            self.sick_leave_taken = sick_leave_taken
        if no_paid_leave_taken is not None:
            self.no_paid_leave_taken = no_paid_leave_taken
        if resigning_date is not None:
            self.resigning_date = resigning_date
        if resigning is not None:
            self.resigning = resigning
        if annual_leave is not None:
            self.annual_leave = annual_leave
        if paid_date is not None:
            self.paid_date = paid_date
        if join_date is not None:
            self.join_date = join_date
        self.curr_date = datetime.date.today()
        first = self.curr_date.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        self.last_month = datetime.date(last_month.year, last_month.month, 1)
    

    def find_probation(self):
        day, month, year = self.join_date.day, self.join_date.month, self.join_date.year
        month += 3
        if month > 12:
            month -= 12
            year += 1

        last_day = calendar.monthrange(year, month)[1]
        if day > last_day:
            day -= last_day
            month += 1
            if month > 12:
                month -= 12
                year += 1
        probation_due = datetime.date(year, month, day)
        return probation_due
    

    def find_leave_taken_deduction(self, probation_due):
        number_of_days = calendar.monthrange(self.curr_date.year, self.last_month.month)[1] 
        accurate_day_salary = self.basic_salary / number_of_days
        total_deduction, paid_leave_salary = 0, 0

        temp = datetime.date(self.paid_date.year, self.paid_date.month, self.paid_date.day)
        remind = -1

        if self.last_month >= probation_due or self.last_month < probation_due < temp:
            if self.sick_leave_taken > 0:
                total_deduction += accurate_day_salary * (self.sick_leave_taken - 1) if self.sick_leave_taken - 1 > 0 else 0
                paid_leave_salary += accurate_day_salary if self.sick_leave_taken >= 1 else accurate_day_salary * 0.5
            total_deduction += accurate_day_salary * self.no_paid_leave_taken
            if self.last_month < probation_due < temp:  
                remind = 1
        elif temp <= probation_due:
            total_leave = self.sick_leave_taken + self.no_paid_leave_taken
            total_deduction += (accurate_day_salary * total_leave)
            remind = 0
        return total_deduction, remind
    

    def find_resign_deduction(self):

        if self.resigning:
            number_of_days = calendar.monthrange(self.curr_date.year, self.last_month.month)[1] 
            accurate_day_salary = self.basic_salary / number_of_days
            if self.resigning_date.year == self.join_date.year:
                start_date = self.join_date
            else:
                start_date = datetime.datetime(self.resigning_date.year, 1, 1)
            date_delta = self.resigning_date - start_date
            annual_limit = round((date_delta.days + 1) / 365 * 10, 1)
            if self.annual_leave <= annual_limit:
                alternate = round((annual_limit - self.annual_leave) * accurate_day_salary, 2)
                return 0, alternate
            else:
                extra_days = self.annual_leave - annual_limit
                resign_deduction = extra_days * accurate_day_salary
                return resign_deduction, 0
        return 0, 0
    

    def MPF_calculation(self, alternate, total_deduction):

        days_at_work = calendar.monthrange(self.curr_date.year, self.last_month.month)[1] if (
                    self.join_date.month != self.curr_date.month - 1 or
                    (self.join_date.month == self.last_month.month and self.join_date.year != self.curr_date.year)) else \
        calendar.monthrange(self.curr_date.year, self.last_month.month)[1] - self.join_date.day + 1
        number_of_days = calendar.monthrange(self.curr_date.year, self.last_month.month)[1] 
        accurate_day_salary = self.basic_salary / number_of_days

        temp_net_payment = accurate_day_salary * days_at_work + alternate - total_deduction
        Employer_MPF_Contribution = round(temp_net_payment * 0.05, 2) if self.basic_salary < 30000 else 1500
        Employee_MPF_Contribution = 0
        date_diff = self.paid_date - self.join_date
        if date_diff.days + 1 >= 60 and temp_net_payment >= 7100:
            Employee_MPF_Contribution = round(temp_net_payment * 0.05,
                                              2) if self.basic_salary < 30000 else 1500
        return Employee_MPF_Contribution, Employer_MPF_Contribution


    def find_final_net_payment(self, alternate, total_deduction):

        days_at_work = calendar.monthrange(self.curr_date.year, self.last_month.month)[1] if (
                    self.join_date.month != self.curr_date.month - 1 or
                    (self.join_date.month == self.last_month.month and self.join_date.year != self.curr_date.year)) else \
        calendar.monthrange(self.curr_date.year, self.last_month.month)[1] - self.join_date.day + 1
        number_of_days = calendar.monthrange(self.curr_date.year, self.last_month.month)[1] 
        accurate_day_salary = self.basic_salary / number_of_days

        net_paid = round(accurate_day_salary * days_at_work + alternate - total_deduction, 2)
        return net_paid

    