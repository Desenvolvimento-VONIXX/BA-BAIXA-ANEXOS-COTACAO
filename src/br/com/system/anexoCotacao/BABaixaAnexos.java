package br.com.system.anexoCotacao;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import com.sankhya.util.JdbcUtils;
import com.sankhya.util.SessionFile;
import com.sankhya.util.UIDGenerator;

import br.com.sankhya.extensions.actionbutton.AcaoRotinaJava;
import br.com.sankhya.extensions.actionbutton.ContextoAcao;
import br.com.sankhya.extensions.actionbutton.Registro;
import br.com.sankhya.jape.EntityFacade;
import br.com.sankhya.jape.core.JapeSession;
import br.com.sankhya.jape.core.JapeSession.SessionHandle;
import br.com.sankhya.jape.dao.JdbcWrapper;
import br.com.sankhya.jape.sql.NativeSql;
import br.com.sankhya.modelcore.util.EntityFacadeFactory;
import br.com.sankhya.ws.ServiceContext;

public class BABaixaAnexos implements AcaoRotinaJava {

	@Override
	public void doAction(ContextoAcao ctx) throws Exception {
		SessionHandle hnd = null;
		Registro[] registros = ctx.getLinhas();

		if (registros.length > 1) {
			ctx.setMensagemRetorno("Selecione apenas um Registro por vez!");
			return;
		} else if (registros.length == 0) {
			ctx.setMensagemRetorno("Selecione um Registro antes de continuar!");
			return;
		}

		Registro registro = registros[0];
		BigDecimal nroCotacao = (BigDecimal) registro.getCampo("NUMCOTACAO");
		String chave = "anexosCotacao_" + UIDGenerator.getNextID();

		try {
			hnd = JapeSession.open();
			List<FileData> anexos = getAnexos(nroCotacao);

			if (anexos.isEmpty()) {
				ctx.setMensagemRetorno("Nenhum anexo encontrado para a cotação " + nroCotacao);
				return;
			}

			// Compacta todos os anexos em um arquivo ZIP
			byte[] zipFile = createZipFile(anexos);
			SessionFile sessionFile = SessionFile.createSessionFile(chave + ".zip", "application/zip", zipFile);
			ServiceContext.getCurrent().putHttpSessionAttribute(chave, sessionFile);

		} catch (Exception e) {
			ctx.mostraErro("Erro ao tentar baixar os anexos: " + e.getMessage());
		} finally {
			JapeSession.close(hnd);
		}

		// Link de download do arquivo ZIP
		ctx.setMensagemRetorno("<p><strong>Para baixar os anexos compactados:</strong></p>"
				+ "<a id=\"alink\" href=\"/mge/visualizadorArquivos.mge?chaveArquivo=" + chave
				+ "&forcarDownload=S\" target=\"_blank\">Baixar Anexos (.zip)</a>");
	}

	private static String detectFileExtension(byte[] content) {
	    if (content.length >= 4) {
	        if (content[0] == 0x25 && content[1] == 0x50 && content[2] == 0x44 && content[3] == 0x46) {
	            return "pdf"; // PDF
	        } else if (content[0] == (byte) 0xFF && content[1] == (byte) 0xD8 && content[2] == (byte) 0xFF) {
	            return "jpg"; // JPEG
	        } else if (content[0] == (byte) 0x89 && content[1] == 0x50 && content[2] == 0x4E && content[3] == 0x47) {
	            return "png"; // PNG
	        } else if (content[0] == (byte) 0xD0 && content[1] == (byte) 0xCF && content[2] == (byte) 0x11
	                && content[3] == (byte) 0xE0) {
	            // Diferenciação entre DOC e XLS
	            if (content.length >= 5 && content[4] == (byte) 0x01) {
	                return "doc"; // DOC (Word 97-2003)
				} else {
	                return "xls"; // XLS (Excel 97-2003)
	            }
	        }
	        // Verificação para XLSX e DOCX (baseados em ZIP)
	        else if (content[0] == (byte) 0x50 && content[1] == (byte) 0x4B && content[2] == (byte) 0x03
	                && content[3] == (byte) 0x04) {
	            // Verificar se é DOCX ou XLSX com base em arquivos internos
	            if (isDocx(content)) {
	                return "docx"; // DOCX
	            } else if (isXlsx(content)) {
	                return "xlsx"; // XLSX
	            }
	        }
	    }
	    // Se não for identificado, retorna binário
	    return "bin";
	}

	// Função para verificar se é um arquivo .docx
	private static boolean isDocx(byte[] content) {
		try (java.util.zip.ZipInputStream zipStream = new java.util.zip.ZipInputStream(
				new java.io.ByteArrayInputStream(content))) {
			// Procurando por "word/document.xml" dentro de um arquivo .docx
			java.util.zip.ZipEntry entry;
			while ((entry = zipStream.getNextEntry()) != null) {
				if (entry.getName().equals("word/document.xml")) {
					return true; // Se encontrar "word/document.xml", é um arquivo .docx
				}
			}
		} catch (IOException e) {
			return false;
		}
		return false;
	}

	// Função para verificar se é um arquivo .xlsx
	private static boolean isXlsx(byte[] content) {
		try (java.util.zip.ZipInputStream zipStream = new java.util.zip.ZipInputStream(
				new java.io.ByteArrayInputStream(content))) {
			// Procurando por "xl/workbook.xml" dentro de um arquivo .xlsx
			java.util.zip.ZipEntry entry;
			while ((entry = zipStream.getNextEntry()) != null) {
				if (entry.getName().equals("xl/workbook.xml")) {
					return true; // Se encontrar "xl/workbook.xml", é um arquivo .xlsx
				}
			}
		} catch (IOException e) {
			return false;
		}
		return false;
	}

	// Classe para armazenar as informações de cada anexo (nome e conteúdo)
	private static class FileData {
		private final String fileName;
		private final byte[] content;

		public FileData(int index, byte[] content) {
			this.content = content;
			String extension = detectFileExtension(content);
			this.fileName = "arquivo_" + index + "." + extension;
		}

		public String getFileName() {
			return fileName;
		}

		public byte[] getContent() {
			return content;
		}
	}

	// Recupera os anexos da cotação, agora com o nome baseado na extensão detectada
	private List<FileData> getAnexos(BigDecimal nroCotacao) throws Exception {
		JdbcWrapper jdbc = null;
		NativeSql sql = null;
		ResultSet rset = null;

		List<FileData> anexosList = new ArrayList<>();
		int index = 1;

		try {
			final EntityFacade entity = EntityFacadeFactory.getDWFFacade();
			jdbc = entity.getJdbcWrapper();
			jdbc.openSession();

			sql = new NativeSql(jdbc);
			sql.appendSql("SELECT ATA.CONTEUDO AS CONTEUDO " + "FROM SANKHYA.TSIATA ATA "
					+ "INNER JOIN SANKHYA.TGFCAB CAB ON CAB.NUNOTA = ATA.CODATA "
					+ "WHERE CAB.NUMCOTACAO = :NROCOTACAO");
			sql.setNamedParameter("NROCOTACAO", nroCotacao);
			rset = sql.executeQuery();

			while (rset.next()) {
				byte[] conteudo = rset.getBytes("CONTEUDO");
				if (conteudo != null) {
					anexosList.add(new FileData(index++, conteudo));
				}
			}

		} catch (Exception e) {
			throw new Exception("Erro ao recuperar anexos: " + e.getMessage(), e);
		} finally {
			JdbcUtils.closeResultSet(rset);
			NativeSql.releaseResources(sql);
			JdbcWrapper.closeSession(jdbc);
		}

		return anexosList;
	}

	// Cria o arquivo ZIP com os anexos
	private byte[] createZipFile(List<FileData> anexos) throws IOException {
		try (ByteArrayOutputStream baos = new ByteArrayOutputStream();
				ZipOutputStream zos = new ZipOutputStream(baos)) {

			for (FileData anexo : anexos) {
				ZipEntry entry = new ZipEntry(anexo.getFileName());
				zos.putNextEntry(entry);
				zos.write(anexo.getContent());
				zos.closeEntry();
			}
			zos.finish();
			return baos.toByteArray();
		}
	}
}
