import os.path
import fdb
import pandas as pd

from config import filepath, addr_from, addr_to, addr_cc, date_yesterday, date_end, date, server
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

connect = fdb.connect(dsn='', user='',
                      password='')

cursor = connect.cursor()

cursor.execute(f"""select dep_name as "Отдел",user_fam as "Фамилия",user_name as "Имя",user_sname as "Отчество",
                    dt_min as "Первый проход",dt_max as "Последний проход",
                    cast(t_rab/60 as varchar(2)) || ':' ||
                    case when t_rab - (t_rab/60)*60 < 10 then '0'|| cast(t_rab - (t_rab/60)*60 as varchar(2))
                    else cast(t_rab - (t_rab/60)*60 as varchar(2))
                    end tm_rab
                    from
                    (
                        select tabnum, dep_name, user_fam, user_name, user_sname, dt_min, dt_max,
                        sum (extract(hour from dt_max) - extract(hour from dt_min))*60 + extract(minute from dt_max)
                        - extract(minute from dt_min) t_rab
                        from
                        (
                            select tabnum,dep_name,user_fam,user_name,user_sname,
                            (
                                select min(event_date) from event_log a1
                                    where a1.user_id = a2.user_id and
                                    extract(day from event_date)=a2.dt and
                                    extract(month from event_date)=a2.mt and
                                    extract(year from event_date)=a2.yr
                            ) dt_min,
                            (
                                select max(event_date) from event_log a1
                                    where a1.user_id = a2.user_id and
                                    extract(day from event_date)=a2.dt and
                                    extract(month from event_date)=a2.mt and
                                    extract(year from event_date)=a2.yr
                            ) dt_max
                            from
                            (
                                Select distinct c.user_id,
                                c.tabnum,
                                c.user_fam,
                                c.user_name,
                                c.user_sname,
                                f.dep_name,
                                extract(day from h.event_date) dt,
                                extract(month from h.event_date) mt,
                                extract(year from h.event_date) yr
                                from access_rights a
                                left outer join TOKENS b on b.token_id = a.token_id
                                left outer join users c on c.user_id = b.user_id
                                inner join departments f on f.dep_id = c.dep_id
                                left outer join event_log h on h.token_id = b.token_id

                               where
                                h.event_date between '{date_yesterday}' and '{date_end}'
                            ) a2
                        )
                        group by tabnum, dep_name, user_fam, user_name, user_sname, dt_min, dt_max
                    )
                    order by dep_name, tabnum, user_fam, dt_min""")

next_row = cursor.fetchall()
cursor.close()
connect.close()


def create_file(filepath):
    df = pd.DataFrame(next_row, columns=['Отдел', 'Фамилия', 'Имя', 'Отчество', 'Первый проход', 'Последний проход',
                                         'TM_RAB'])
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as wb:
        df.to_excel(wb, sheet_name='Учет рабочего времени', index=False)
        sheet = wb.sheets['Учет рабочего времени']
        sheet.autofilter(0, 0, df.shape[0], 2)
        for col_idx, col_name in enumerate(df.columns):
            max_width = max([len(col_name)] + [len(str(s)) for s in df[col_name]])
            sheet.set_column(col_idx, col_idx, max_width + 10)

    filename = os.path.basename(filepath)
    ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)

    with open(filepath, 'rb') as fp:
        file = MIMEBase(maintype, subtype)
        file.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(file)
        file.add_header('Content-Disposition', 'attachment', filename=filename)

    return file


def send_mail(addr_from, addr_to, file):  # , addr_cc):
    msg = MIMEMultipart()
    msg['From'] = addr_from
    msg['To'] = ", ".join(addr_to)  # addr_to
    # msg['CC'] = ", ".join(addr_cc)
    msg['Subject'] = f'Отчет об опоздавших за {date.today().strftime("%b-%d-%Y")}'

    if os.path.isfile(filepath):
        msg.attach(file)

    server.starttls()
    server.send_message(msg)
    server.quit()


if __name__ == "__main__":
    file = create_file(filepath)
    send_mail(addr_from, addr_to, file)  # , addr_cc)
