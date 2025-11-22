import { GraphFI } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/photos';

// Serviço para obter fotos de perfil via Graph API
export class GraphService {
  private _graph: GraphFI;

  // Cache para não repetir chamadas
  private _fotoCache: Map<string, string> = new Map();

  constructor(graph: GraphFI) {
    this._graph = graph;
  }

  // Obtém a foto de perfil de um utilizador
  public async getFotoUtilizador(email: string): Promise<string | undefined> {
    // Verifica se já está em cache
    if (this._fotoCache.has(email)) {
      return this._fotoCache.get(email);
    }

    try {
      // Obtém a foto do utilizador via Graph API
      const photoBlob = await this._graph.users
        .getById(email)
        .photo
        .getBlob();

      // Converte o blob para base64
      const fotoUrl = await this.blobToBase64(photoBlob);

      // Guarda em cache
      this._fotoCache.set(email, fotoUrl);

      return fotoUrl;
    } catch (error) {
      // Se o utilizador não tiver foto, retorna undefined
      // Não é necessariamente um erro crítico
      console.warn(`Foto não encontrada para ${email}:`, error);
      return undefined;
    }
  }

  // Obtém as fotos de vários utilizadores em paralelo
  public async getFotosUtilizadores(
    emails: string[]
  ): Promise<Map<string, string>> {
    const resultados = new Map<string, string>();

    // Filtra emails únicos e não vazios
    const emailsUnicos = Array.from(new Set(emails.filter(e => e)));

    // Obtém as fotos em paralelo
    const promessas = emailsUnicos.map(async email => {
      const fotoUrl = await this.getFotoUtilizador(email);
      if (fotoUrl) {
        resultados.set(email, fotoUrl);
      }
    });

    await Promise.all(promessas);

    return resultados;
  }

  // Converte blob para base64
  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        resolve(reader.result as string);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }
}

export default GraphService;
