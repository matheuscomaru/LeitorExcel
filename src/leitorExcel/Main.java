package leitorExcel;

import java.io.IOException;
import java.util.ArrayList;

public class Main {

	public static void main(String[] args) {

		LeitorArquivo ger = new LeitorArquivo();

		System.out.println("LENDO EXCEL");

		ArrayList<Cliente> lista = new ArrayList<>();

		try {
			lista = ger.leitorNomes();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		System.out.println("Quantidade de clientes :" + lista.size());

		for (Cliente cliente : lista) {
			System.out.println("---------------------");
			System.out.println("Codigo : " + cliente.getCodigo());
			System.out.println("Razão : " + cliente.getRazaoSocial());
			System.out.println("Fantasia : " + cliente.getFantasia());
			System.out.println("Endereço : " + cliente.getEndereco());

		}

	}

}
