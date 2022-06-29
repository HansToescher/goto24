package at.goto24.data.util;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.Properties;

import org.apache.log4j.Logger;

public class ClassFactory {
    private static final Logger LOG = Logger.getLogger(ClassFactory.class);

    private Properties prop = new Properties();

    public ClassFactory(InputStream stream) throws IOException {
	prop.load(stream);
    }

    public Object createObject(String name) throws ClassNotFoundException, InstantiationException, IllegalAccessException {

	String classname = prop.getProperty(name);

	if (classname != null) {
	    Class<?> cl = Class.forName(classname);

	    return cl.newInstance();
	}

	return null;
    }

    public Object createObjectWithConstructor(String name, Class<?> cparamcl, Object obj) throws ClassNotFoundException, NoSuchMethodException, InstantiationException, IllegalAccessException, InvocationTargetException {
	String classname = prop.getProperty(name);

	if (classname != null) {
	    Class<?> cl = Class.forName(classname);
	    Constructor<?> cons = cl.getDeclaredConstructor(cparamcl);

	    return cons.newInstance(obj);
	} else {
	    LOG.debug("Type not Found:" + name);
	}

	return null;
    }

    public Object createObjectWithConstructor(String name, Class<?> cparamcl1, Object obj1, Class<?> cparamcl2, Object obj2, Class<?> cparamcl3, Object obj3) throws ClassNotFoundException, NoSuchMethodException, InstantiationException, IllegalAccessException, InvocationTargetException {
	String classname = prop.getProperty(name);

	if (classname != null) {
	    Class<?> cl = Class.forName(classname);
	    Constructor<?> cons = cl.getDeclaredConstructor(cparamcl1, cparamcl2, cparamcl3);

	    return cons.newInstance(obj1, obj2, obj3);
	} else {
	    LOG.debug("Type not Found:" + name);
	}

	return null;
    }
}
