#!/bin/sh

if [ -z "$1" ]
then
  echo "Usage: $0 <Role arn:aws:iam::....:....>"
  exit 1
fi

NAME=$(basename $(realpath .))
cd src && \
    zip -r ../${NAME}.zip . && \
    cd .. && \
    aws lambda create-function --function-name ${NAME} --zip-file fileb://${NAME}.zip --handler index.handler --runtime nodejs10.x --role $1 --timeout 30&& \
    aws lambda create-alias --function-name ${NAME} --name service --function-version \$LATEST
