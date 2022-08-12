# VbaLinkBudBackup

This file is for design and management of link budget of Indoor distributed control system (DCS)  network. The aim for this project is to automate the tedious calculations in designing routing of feeder cable and choice of coupler in the whole rf network.

## Content

1. [Introduction](##Introduction)
2. [User Manual](##User-Manual)

\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\
\

## Introduction
The main problem in indoor DCS system is the complexity in calculation to reach a balance of power output in every antenna. Similar to every distrubuted system (e.g. Electiricty, Water distrubuted system), the system has a source to distrubute to every descending separation point (couplers) and eventually reach the 
output (antenna). The complexity of DCS system design comes from the limitation in construction site, mainly the pathway restricted and the antenna location pre-requested in tender by clients. To ensure the calculations reach the demanded power output (RSRP), link budget calculation is required. Without automation, 
we need to care about the choice of couplers, length and routing of feeder cables. Changing a coupler will alter the descendents' output which consisted in the same routing, which makes link budgets design difficult and complex. 

This project focus on calculation of RSRP output from antenna by subtracting the coupling loss and feeder loss in the whole pathway from source to ouput antenna. Visio is ultilized as GUI for showcase of link budget and DCS network design. 

## User Manual


