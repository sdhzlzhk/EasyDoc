package com.glodon.hwpf;

import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author liuzk
 * @create 2016-12-21 15:12.
 */
public class DocPictureManager implements PicturesManager {
    private static final String IMG_RELATIVE_DIR = "img/";
    @Override
    public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
        return null;
    }

    public String savePicture(Picture picture) throws IOException {
        String imgPath = System.getProperty("user.dir") + IMG_RELATIVE_DIR + picture.suggestFullFileName();
        File imgOut = new File(imgPath);
        if(!imgOut.exists()){
            imgOut.getParentFile().mkdirs();
        }
        picture.writeImageContent(new FileOutputStream(imgPath));
        return imgOut.getCanonicalPath();
    }
}
