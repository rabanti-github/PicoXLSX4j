/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.lowLevel;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.*;
import org.w3c.dom.Document;

/**
 * Class representing packer to compile XLSX
 * @author Raphael Stoeckli
 */
public class Packer {
    
    /**
     * Path of the main content type file (MSXML)
     */
    static final String CONTENTTYPE_DOCUMENT = "[Content_Types].xml";
    
    private List<byte[]> dataList;
    private List<String> pathList;
    private List<String> contentTypeList;
    private List<Relationship> relationships;
    private List<Boolean> includeContentType;
    
    /**
     * Default constructor
     */
    public Packer()
    {
        dataList = new ArrayList<>();
        pathList = new ArrayList<>();
        contentTypeList = new ArrayList<>();
        relationships = new ArrayList<>();
        includeContentType = new ArrayList<>();
    }
    
    /**
     * Method to pack the file into a XLSX file. This is the actual compiling and saving method
     * @param fileName Filename of the XLSX file
     * @throws picoxlsx4j.exception.IOException Thrown in case of a error while saving
     */
    public void pack(String fileName) throws picoxlsx4j.exception.IOException
    {
        try
        {
        byte[] contentTypes = createContenTypeDocument();
        FileOutputStream dest = new FileOutputStream(fileName);
        ZipOutputStream out = new ZipOutputStream(new BufferedOutputStream(dest), Charset.forName("UTF-8"));
        out.setMethod(ZipOutputStream.DEFLATED);
        ZipEntry entry = new ZipEntry(CONTENTTYPE_DOCUMENT);
        out.putNextEntry(entry);
        out.write(contentTypes, 0, contentTypes.length);
        byte[] data;
        Document doc;
        for (int i = 0; i < this.relationships.size(); i++)
        {
            data = createRelationshipDocument(this.relationships.get(i));
            entry = new ZipEntry(this.relationships.get(i).getRootFolder());
            out.putNextEntry(entry);
            out.write(data, 0, data.length);
        }
        for (int i = 0; i < this.dataList.size(); i++)
        {
            data = this.dataList.get(i);
            entry = new ZipEntry(this.pathList.get(i));
            out.putNextEntry(entry);
            out.write(data, 0, data.length);
        }
        out.flush();
        out.close();
        }
        catch(Exception e)
        {
            throw new picoxlsx4j.exception.IOException("There was an error while packing the file. Please see the inner exception.", e);
        }
    }    
    
    /**
     * Adds a Part to the file
     * @param name Filename with relative path
     * @param contentType URL with information about the content type (MSXML).<br>This information is used in the main content type file
     * @param document XML document to add
     * @throws picoxlsx4j.exception.IOException Thrown if the document could not be converted to a byte array
     */
    public void addPart(String name, String contentType, Document document) throws picoxlsx4j.exception.IOException
    {
        dataList.add(LowLevel.createBytesFromDocument(document));
        pathList.add(name);
        contentTypeList.add(contentType);
        includeContentType.add(true);
    }
    
    /**
     * Adds a Part to the file
     * @param name Filename with relative path
     * @param contentType URL with information about the content type (MSXML).<br>This information is used in the main content type file
     * @param document XML document to add
     * @param includeInContentType If true, the content type will be added in the main content type file, otherwise not
     * @throws picoxlsx4j.exception.IOException Thrown if the document could not be converted to a byte array
     */
    public void addPart(String name, String contentType, Document document, boolean includeInContentType) throws picoxlsx4j.exception.IOException
    {
        dataList.add(LowLevel.createBytesFromDocument(document));
        pathList.add(name);
        contentTypeList.add(contentType);
        includeContentType.add(includeInContentType);
    }    
    
    /**
     * Creates a relationship. This will be used to generate a .rels file in the compilation (MSXML)
     * @param path relative path and filename to rels file (e.g. _rels/.rels)
     * @return Returns the object reference to add relationship entries
     */
    public Relationship createRelationship(String path)
    {
        Relationship r = new Relationship(path);
        this.relationships.add(r);
        return r;
    }
    
    /**
     * Creates the main content type file (MSXML)
     * @return Returns the byte array to add into the compilation
     * @throws picoxlsx4j.exception.IOException Thrown if the document could not be converted to a byte array
     */
    private byte[] createContenTypeDocument() throws picoxlsx4j.exception.IOException
    {
        StringBuilder sb = new StringBuilder();
        sb.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n");
        sb.append("<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />\r\n");        
        sb.append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />\r\n");

        for (int i = 0; i < this.contentTypeList.size(); i++)
        {
            if (this.includeContentType.get(i) == false) { continue; }
            sb.append("<Override PartName=\"/");
            sb.append(this.pathList.get(i));
            sb.append("\" ContentType=\"");
            sb.append(this.contentTypeList.get(i));
            sb.append("\" />\r\n");
        }
        sb.append("</Types>");
        Document d = LowLevel.createXMLDocument(sb.toString());
        return LowLevel.createBytesFromDocument(d);
    }
    
    /**
     * Creates a relationship file (MSXML)
     * @param rel Relationship object to process
     * @return  Returns the byte array to add into the compilation
     * @throws picoxlsx4j.exception.IOException Thrown if the document could not be converted to a byte array
     */
    private byte[] createRelationshipDocument(Relationship rel) throws picoxlsx4j.exception.IOException
    {
        StringBuilder sb = new StringBuilder();
        sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n");
        for (int i = 0; i < rel.getIdList().size(); i++)
        {
            sb.append("<Relationship Target=\"");
            sb.append(rel.getTargetList().get(i));
            sb.append("\" Type=\"");
            sb.append(rel.getTypeList().get(i));
            sb.append("\" Id=\"");
            sb.append(rel.getIdList().get(i));
            sb.append("\"/>\r\n");
        }
        sb.append("</Relationships>");
        Document d = LowLevel.createXMLDocument(sb.toString());
        return LowLevel.createBytesFromDocument(d);
    }    
    
    /**
     * Nested class representing a relationship (MSXML)
     */
    public class Relationship
    {
        private String rootFolder;
        private List<String> targetList;
        private List<String> typeList;
        private List<String> idList;
        private int currentId;

        /**
         * Gets the root folder of the relationship
         * @return Root folder of the relationship
         */
        public String getRootFolder() {
            return rootFolder;
        }

        /**
         * Gets the list of targets
         * @return ArrayList of targets as string
         */
        public List<String> getTargetList() {
            return targetList;
        }

        /**
         * Gets the list of types
         * @return ArrayList of types as string
         */
        public List<String> getTypeList() {
            return typeList;
        }

        /**
         * Gets the list of IDs (rId...)
         * @return ArrayList of IDs as string
         */
        public List<String> getIdList() {
            return idList;
        }
        
        /**
         * Constructor with definition of the root folder
         * @param path Root folder of the relationship
         */
        public Relationship(String path)
        {
            this.idList = new ArrayList<>();
            this.targetList = new ArrayList<>();
            this.typeList = new ArrayList<>();
            this.rootFolder = path;
            this.currentId = 1;
        }
        
        /**
         * Adds a relationship entry to the relationship
         * @param target Target of the entry
         * @param type Type of the entry 
         */
        public void addRelationshipEntry(String target, String type)
        {
            this.targetList.add(target);
            this.typeList.add(type);
            String id = "rId" + Integer.toString(this.currentId);
            this.idList.add(id);
            this.currentId++;
        }   
    }
    
}
