package com.glodon.tika;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;

/**
 * Created by zhongkai on 2016/12/18.
 */
public class SafeWordExtractor {

    private HWPFDocument document;

    public SafeWordExtractor(InputStream inputStream) throws IOException {
        this(HWPFDocument.verifyAndBuildPOIFS(inputStream));
    }
    private SafeWordExtractor(POIFSFileSystem fs) throws IOException {
        this(new HWPFDocument(fs));
    }

    private SafeWordExtractor(HWPFDocument document) {
        this.document = document;
    }
}
