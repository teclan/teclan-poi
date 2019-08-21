package com.teclan.word;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DefaultWordTextHandler implements WordTextHandler {
    private static final Logger LOGGER = LoggerFactory.getLogger(DefaultWordTableHandler.class);

    public void handler(String text) {
        LOGGER.info("{}",text);
    }
}
