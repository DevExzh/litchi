import com.sun.star.beans.PropertyValue;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.document.MacroExecMode;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.io.File;

/** Stores one already-openable ODB through a caller-provided LibreOffice UNO endpoint. */
public final class UnoStore {
    private UnoStore() {}

    public static void main(String[] arguments) throws Exception {
        if (arguments.length != 2) {
            throw new IllegalArgumentException("expected UNO URL and one ODB path");
        }
        XComponentContext localContext = Bootstrap.createInitialComponentContext(null);
        Object resolverService = localContext.getServiceManager().createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        );
        XUnoUrlResolver resolver = UnoRuntime.queryInterface(
            XUnoUrlResolver.class, resolverService
        );
        Object remoteContext = resolver.resolve(arguments[0]);
        XComponentContext context = UnoRuntime.queryInterface(
            XComponentContext.class, remoteContext
        );
        if (context == null) {
            throw new IllegalStateException("UNO URL did not resolve a component context");
        }
        Object desktop = context.getServiceManager().createInstanceWithContext(
            "com.sun.star.frame.Desktop", context
        );
        XComponentLoader loader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);
        PropertyValue hidden = new PropertyValue();
        hidden.Name = "Hidden";
        hidden.Value = Boolean.TRUE;
        PropertyValue macros = new PropertyValue();
        macros.Name = "MacroExecutionMode";
        macros.Value = Short.valueOf(MacroExecMode.NEVER_EXECUTE);
        String url = new File(arguments[1]).getCanonicalFile().toURI().toString();
        XComponent component = loader.loadComponentFromURL(
            url, "_blank", 0, new PropertyValue[] {hidden, macros}
        );
        if (component == null) {
            throw new IllegalStateException("LibreOffice did not load the ODB component");
        }
        try {
            XStorable storable = UnoRuntime.queryInterface(XStorable.class, component);
            if (storable == null) {
                throw new IllegalStateException("ODB component has no XStorable facade");
            }
            storable.store();
        } finally {
            component.dispose();
        }
        System.out.println(url);
    }
}
