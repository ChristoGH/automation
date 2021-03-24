library(shiny)
runExample("01_hello")



path="C:\\Users\\christo.strydom\\github_repos\\automation\\CBR\\"
setwd(path)

mip_claims='mip_claims.csv'
mip_members='mip_members.csv'
owls_claims='owls_claims.csv' # 
owls_members='owls_members.csv'
owls_terminations='terminations_raw.csv'
owls_received_claims='received_claims.csv'


owls_claims.df <- read.csv(file = owls_claims)
