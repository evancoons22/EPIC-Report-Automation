import matplotlib.pyplot as plt
import matplotlib
import numpy as np
import io
import base64
import seaborn as sns
import pandas as pd


#importing data
# portfolio = pd.read_excel("Epic Portfolio.xlsx")
# reporting_list = pd.read_excel("reportinglist1.xlsx")
# commitments_and_transactions = pd.read_excel("commitmentstransactions1.xlsx")

def build_report(address, portfolio, reporting_list, commitments_and_transactions, count):
  print(f"report {count} building...")
  # portfolio = pd.read_excel("Epic Portfolio.xlsx")
  # reporting_list = pd.read_excel("reportinglist1.xlsx")
  # commitments_and_transactions = pd.read_excel("commitmentstransactions1.xlsx")
# manipulating, cleaning, joining
  commitments = commitments_and_transactions.loc[commitments_and_transactions["Action"] == "Commitment"]
  commitments.loc[commitments["Fund Name"] == "EPIC-M7 LP", "Class"] = "Class M"
  reporting_list = reporting_list.rename(columns = {"Account Name": "Investor"})
  investorreport = commitments.merge(reporting_list, how = "inner", on = "Investor")
  portfolio = portfolio.rename(columns = {"Class-combined": "Class"})
  investorreport = investorreport.merge(portfolio, how = "inner", on ="Class")[["Email", "Investor", "Class", "Amount", "Date", "EPIC Investment Vehicle", "Share Class", "First"]].drop_duplicates()

  address_report = investorreport.loc[investorreport["Email"] == address]

  #creating pivot table from email
  address_plot1 = address_report.pivot_table(index = 'Investor', values = 'Amount', aggfunc = 'sum', columns = 'Class')

  #plot 1
  ax = address_plot1.plot(kind = 'bar', stacked = True, figsize = (10,8))
  plt.legend(["EPIC ACRE", "Fund II Class E", "Fund II Class I", "Fund I Class E", "Fund I Class I", "EPIC M7"])
  ax.get_yaxis().set_major_formatter(
      matplotlib.ticker.FuncFormatter(lambda x, p: '$' + format(int(x), ',')))
  ax.grid(axis = 'y', which = 'major', linestyle = 'dashed')
  plt.tight_layout()

  plot1 = io.BytesIO()
  plt.savefig(plot1, format='jpg')
  plot1.seek(0)
  plot1_data = f"data:image/png;base64,{base64.b64encode(plot1.read()).decode()}"

  #plot 2 creating second pivot table and second plot
  address_plot2 = address_report.pivot_table(index = "EPIC Investment Vehicle", values = "Amount", aggfunc = "sum")
  ax2 = address_plot2.plot(kind = 'bar', stacked = True, figsize = (10,8))
  ax2.get_yaxis().set_major_formatter(
      matplotlib.ticker.FuncFormatter(lambda x, p: '$' + format(int(x), ',')))
  ax2.grid(axis = 'y', which = 'major', linestyle = 'dashed')
  plt.xticks(rotation = 0)
  plt.title("Your Invesments by Fund")

  plot2 = io.BytesIO()
  plt.savefig(plot2, format='jpg')
  plot2.seek(0)
  plot2_data = f"data:image/png;base64,{base64.b64encode(plot2.read()).decode()}"

  #plot 3 
  address_plot3 = address_report[["Amount", "Date"]].copy(deep = True)
  address_plot3["Date"] = address_plot3["Date"].apply(lambda x: x.year)
  address_plot3 = address_plot3.set_index("Date", drop = True)
  ax3 = address_plot3.pivot_table(index = 'Date', values = 'Amount', aggfunc = 'sum').plot(kind = 'bar', figsize = (10,8))
  ax3.get_yaxis().set_major_formatter(
      matplotlib.ticker.FuncFormatter(lambda x, p: '$' + format(int(x), ',')))
  ax3.grid(axis = 'y', which = 'major', linestyle = 'dashed')
  plt.xticks(rotation = 0)
  plt.title("Your Investments by Year")

  plot3 = io.BytesIO()
  plt.savefig(plot3, format='jpg')
  plot3.seek(0)
  plot3_data = f"data:image/png;base64,{base64.b64encode(plot3.read()).decode()}"

  amount = int(sum(address_report['Amount']))
  amount = ("{:,}".format(amount))

  first = ''
  for i in address_report["First"]: 
    if not pd.isnull(i): 
        first = i
        break

  html = f'''
    <!DOCTYPE html>
  <html>
  <head>
  <title>EPIC Report Template</title>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <!-- <link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css"> -->
  <link rel="stylesheet" href="../mainpython.css">
  </head>
  <body>

  <!-- Navbar (sit on top) -->
  <div class="w3-top">
    <div class="w3-bar w3-white w3-wide w3-padding w3-card">
      <a href ="https://www.epic-funds.com/epics-solution/" target="_blank">
        <img class="w3-bar-item" src="../logo.png" alt="logo" width="50" height="50">
      </a>

      <a href="#home" class="w3-bar-item w3-button"><b>EPIC</b> Platform </a>
      <!-- Float links to the right. Hide them on small screens -->
      <div class="w3-right w3-hide-small">
        <a href="#projects" class="w3-bar-item w3-button">Your Portfolio</a>
        <a href="#about" class="w3-bar-item w3-button">Note</a>
        <a href="#contact" class="w3-bar-item w3-button">Contact</a>
      </div>
    </div>
  </div>

  <!-- Header -->
  <header class="w3-display-container w3-content w3-wide" style="max-width:1500px;" id="home">
    <img class="w3-image" src="/w3images/architect.jpg" alt="Architecture" width="1500" height="800">
    <div class="w3-display-middle w3-margin-top w3-center">
      <h1 class="w3-xxlarge w3-text-white"><span class="w3-padding w3-black w3-opacity-min"><b>BR</b></span> <span class="w3-hide-small w3-text-light-grey">Architects</span></h1>
    </div>
  </header>

  <!-- Page content -->
  <div class="w3-content w3-padding" style="max-width:1564px">

    <!-- Project Section -->
    <div class="w3-container w3-padding-32" id="projects">
      <h3 class="w3-border-bottom w3-border-light-grey w3-padding-16">{first + "'s"} Portfolio</h3>
    </div>

    <div class="w3-row-padding">
      <div class="w3-col l3b m6 w3-margin-bottom">
        <div class="w3-display-container">
          <img src="{plot1_data}" alt="House" style="width:100%">
        </div>
      </div>
      <div class="w3-col l3b m6 w3-margin-bottom">
        <div class="w3-display-container">
          <h1 class="textbox">{'$' + str(amount)}</h1>
          <p class = "textbox2"> Total Commitment</p>
        </div>
      </div>
    </div>

    <div class="w3-row-padding">
      <div class="w3-col l3b m6 w3-margin-bottom">
        <div class="w3-display-container">
          <img src="{plot3_data}" alt="House" style="width:99%">
        </div>
      </div>
      <div class="w3-col l3b m6 w3-margin-bottom">
        <div class="w3-display-container">
          <div class="w3-display-topleft w3-white w3-padding"></div>
          <img src="{plot2_data}" alt="House" style="width:99%">
        </div>
      </div>
    </div>

    <!-- About Section -->
    <div class="w3-container w3-padding-32" id="about">
      <h3 class="w3-border-bottom w3-border-light-grey w3-padding-16">This Quarter's Performance</h3>
      <p>Since our last update in December, we've officially closed Fund II with a total of $55 million in commitments, bringing our cumulative AUM across all EPIC vehicles to just over $123 million. We've also made 5 new investments from Fund II, bringing our total commitments from Fund II to 72% and 68% for Class I and Class E, respectively.
  We designed the EPIC strategy with the expectation that we would be deploying capital through various economic cycles. While the war in Ukraine, high inflation, challenged supply chains, and changing monetary policy (read: rising rates) have led to a general risk-off sentiment and sell-offs in public and crypto markets, we have not yet experienced meaningful changes in our portfolios. Commenting on what happens from here would be a guess. That said, there seems to be a shifting tide in the market that is exposing weak business models and areas of excess. The period of easy money (i.e., crypto, NFTs, meme stocks, and a "stocks only go up" mentality) seems to have ended and we are confident in the roster of managers we have on our team to navigate what is to come.
      </p>
    </div>

    <div class="w3-row-padding w3-grayscale">
      <div class="w3-col l3 m6 w3-margin-bottom">
        <img src="../profilepics/Jimmy.jpg" alt="Jimmy" style="width:100%">
        <h3>Jimmy Hirschmann</h3>
        <p class="w3-opacity">Managing Director, Investments</p>
        <p>Jimmy is experienced in portfolio management, private equity, investment banking and sales. He was previously an Associate at Envisage Advisors, focused on lower middle market private equity and investment banking. He also previously served as a Portfolio Management Associate at Aesir Capital Management, a liquid credit focused hedge fund acquired by ExodusPoint Capital Management. Prior to Aesir Capital, Jimmy was as an analyst at UBS on the Institutional Fixed Income Credit Sales desk. Jimmy graduated with a BA in Political Science from University of Southern California.</p>
      </div>
      <div class="w3-col l3 m6 w3-margin-bottom">
        <img src="../profilepics/Alec.jpg" alt="Alec" style="width:100%">
        <h3>Alec Garza</h3>
        <p class="w3-opacity">Managing Principal and Co-Founder</p>
        <p>Alec co-founded EPIC after spending 6 years at IWP Family Office ("IWP"), a multi-billion-dollar multi-family office. Alec's responsibilities at IWP included public and private investment research, portfolio allocation construction and execution, and related family office investment management duties. Alec was instrumental in developing the EPIC investment strategy and due diligence processes. Alec graduated from the University of Denver - Daniels College of Business with a BSBA in Finance and minors in Accounting and Spanish. He holds a Series 65 License.</p>
      </div>
      <div class="w3-col l3 m6 w3-margin-bottom">
        <img src="../profilepics/Michael.jpg" alt="Mike" style="width:100%">
        <h3>Michael Fitzpatrick</h3>
        <p class="w3-opacity">Head of BD & IR</p>
        <p>Michael joined EPIC to lead business development and investor relations after working as an operations lead at Uber Technologies and as a management consultant at PwC Advisory. At Uber, Michael led merchant acquisition for new verticals, where he executed go-to-market strategies that focused on partnering with non-restaurant merchants. At PwC, he specialized in finance operations, helping clients of varying size and sector, to implement new processes, systems, and organizational designs. Michael graduated from the University of California, Los Angeles with a BA in Business Economics and a minor in Spanish.</p>
      </div>
    </div>

    <!-- Contact Section -->
    <div class="w3-container w3-padding-32" id="contact">
      <h3 class="w3-border-bottom w3-border-light-grey w3-padding-16">Contact</h3>
      <p>Lets get in touch and talk about your next project.</p>
      <form action="/action_page.php" target="_blank">
        <input class="w3-input w3-border" type="text" placeholder="Name" required name="Name">
        <input class="w3-input w3-section w3-border" type="text" placeholder="Email" required name="Email">
        <input class="w3-input w3-section w3-border" type="text" placeholder="Subject" required name="Subject">
        <input class="w3-input w3-section w3-border" type="text" placeholder="Comment" required name="Comment">
        <button class="w3-button w3-black w3-section" type="submit">
          <i class="fa fa-paper-plane"></i> SEND MESSAGE
        </button>
      </form>
    </div>
    
  <!-- Image of location/map -->
  <div class="w3-container">
    <img src="/w3images/map.jpg" class="w3-image" style="width:100%">
  </div>

  <!-- End page content -->
  </div>


  <!-- Footer -->
  <footer class="w3-center w3-black w3-padding-16">
    <p>EPIC Platform <a href="https://www.epic-funds.com/" title="W3.CSS" target="_blank" class="w3-hover-text-green">epic-funds.com</a></p>
  </footer>

  </body>
  </html>

      '''
  with open(f'reports/{address}_report.html', 'w') as f:
      f.write(html)
  print(f"report {count} finished")