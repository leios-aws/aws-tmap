#!/bin/sh

if [ -z "$1" ]
then
  echo "Usage: $0 <Role arn:aws:iam::....:....>"
  exit 1
fi

rm -f ${NAME}.zip ${NAME}-modules.zip

NAME=$(basename $(realpath .))
cd src && \
    zip -r ../${NAME}-modules.zip node_modules && \
    zip -r ../${NAME}.zip . -x node_modules\* && \
    cd .. && \
    LAYER_VERSION=$(aws lambda publish-layer-version --layer-name ${NAME}-modules --zip-file fileb://${NAME}-modules.zip | jq -r .LayerVersionArn) && \
    aws lambda create-function --function-name ${NAME} --zip-file fileb://${NAME}.zip --handler index.handler --runtime nodejs10.x --role $1 --timeout 30 --layers ${LAYER_VERSION} --environment Variables="{NODE_PATH=/opt/node_modules}" && \
    aws lambda create-alias --function-name ${NAME} --name service --function-version \$LATEST
