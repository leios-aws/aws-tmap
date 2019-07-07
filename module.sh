#!/bin/sh
NAME=$(basename $(realpath .))
REVISION=$(git describe --all | cut -d '/' -f 2)

rm -f ${NAME}-modules.zip

cd src && \
    zip -r ../${NAME}-modules.zip node_modules && \
    cd .. && \
    LAYER_VERSION=$(aws lambda publish-layer-version --layer-name ${NAME}-modules --zip-file fileb://${NAME}-modules.zip | jq -r .LayerVersionArn) && \
    aws lambda update-function-configuration --function-name ${NAME} --layers ${LAYER_VERSION}
