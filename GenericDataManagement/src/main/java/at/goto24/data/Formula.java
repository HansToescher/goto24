package at.goto24.data;


public class Formula {
    final String formul;
    
    public Formula(String formul) {
        this.formul = formul;
    }

    @Override
    public String toString() {
        return formul;
    }
}
