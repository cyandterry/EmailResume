#!/bin/bash

if [ -d "./Personal_Data" ]; then
    echo "Setup folder already exist so won't overwrite. Please delete the folder ./Personal_Data if you want to re-setup"
else
    mkdir Personal_Data

    cd Personal_Data

    echo "Creating nessasary files in ./Personal_Data"
    echo "Please overwrite them without changing filenames"

    # Gmail account setting
    printf "userName = 'YourUserName@gmail.com'\n" >  gmail_account.txt
    printf "passWord = 'YourPassWord'" >>  gmail_account.txt

    # CV template
    touch CV.html

    # Recipient's information
    touch info_list.xls

    # Personal Resume
    touch Resume.pdf

    # Transcript is an option
    touch Transcript.pdf

    # GRE score is an option
    touch GRE.pdf

    echo "Done!"

fi
