import os
import datetime
import pandas as pd
import matplotlib.pyplot as plt
import smtplib
from email import MIMEMultipart
from email import MIMEText
from email import MIMEBase
from email import encoders


def get_data(path, dates):
    """ Reads the excel file specified by the path and returns a dataframe """
    df = pd.DataFrame(index=dates)
    df_temp = pd.read_excel(path, header=1, index_col='Date')
    df = df.join(df_temp, how='inner')
    df = df.dropna()
    return df


def get_date_range(n_days):
    today = datetime.date.today()
    from_date = today - datetime.timedelta(days=n_days)
    return pd.date_range(from_date, today)


def get_bollinger_bands(values, window):
    """Return upper and lower Bollinger Bands."""
    rm = values.rolling(window=window, center=False).mean()
    rstd = values.rolling(window=window, center=False).std()
    upper_band = rm + 2 * rstd
    lower_band = rm - 2 * rstd
    return upper_band, lower_band, rm


def save_plot(df, upper_band, lower_band, rm, fund_name):
    plt.figure()
    ax = df['NAV'].plot(label=fund_name)
    rm.plot(label='Rolling mean', ax=ax)
    upper_band.plot(label='upper band', ax=ax)
    lower_band.plot(label='lower band', ax=ax)
    # Add axis labels and legend
    ax.set_xlabel("Date")
    ax.set_ylabel("Price")
    ax.legend(loc='upper left')
    plt.savefig("plots/{}.png".format(fund_name))


def email_results(msg_body):
    from_address = "sender.example@mail.com"
    recipients = ['reciever1.example@mail.com', 'reciever2.example@mail.com']
    to_address = ", ".join(recipients)
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = "FundBot"
    msg.attach(MIMEText(msg_body, 'plain'))

    for filename in os.listdir(os.path.join("plots")):
        if filename.endswith(".png"):
            attachment = open(os.path.join("plots", filename), "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_address, "123456789")
    text = msg.as_string()
    server.sendmail(from_address, to_address, text)
    server.quit()


def run_main():
    # use date range of the past 90 days
    dates = get_date_range(90)

    msg_body = ""

    # loop over the files in the data folder
    for filename in os.listdir(os.path.join("data")):
        if filename.endswith(".xls"):
            fund_name = filename.strip('.xls')

            rel_path = os.path.join("data", filename)  # get the relative path of the xls files
            df = get_data(rel_path, dates)
            upper_band, lower_band, rm = get_bollinger_bands(df['NAV'], window=20)

            price_today = df.values[-1][0]
            price_yesterday = df.values[-2][0]

            # BUY signal
            if price_yesterday < lower_band[-2] and price_today > price_yesterday:
                msg_body += "BUY" + fund_name + "\n"
                save_plot(df, upper_band, lower_band, rm, fund_name)

            # SELL signal
            if price_yesterday > upper_band[-2] and price_today < price_yesterday:
                msg_body += "SELL" + fund_name + "\n"
                save_plot(df, upper_band, lower_band, rm, fund_name)

            # for debug: save_plot(df, upper_band, lower_band, rm, fund_name)

    print datetime.date.today()
    if not msg_body:
        print "Nothing to report."
    else:
        email_results(msg_body)


if __name__ == "__main__":
    run_main()
