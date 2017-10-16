#!/bin/bash

#author liubo

nohup java -jar  target/center-1.0-SNAPSHOT.jar  >> out 2>> error.log &
echo $?
#tail -f out
