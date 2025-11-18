// metadata-components.tsx
import * as React from "react";
import { Dropdown, IDropdownOption, ComboBox, IComboBox, IComboBoxOption, Spinner, SpinnerSize } from "office-ui-fabric-react";
import { MetadataComponentsProps, MetadataState } from "../interfaces/interfaces";

export class MetadataComponents extends React.Component<MetadataComponentsProps, MetadataState> {
  private searchTimeoutId: NodeJS.Timeout | null = null;
  private readonly SEARCH_DELAY = 300;
  private readonly MAX_VISIBLE_ITEMS = 100;

  constructor(props: MetadataComponentsProps) {
    super(props);
    this.state = {
      clientes: [],
      clientesFiltered: [],
      selectedCliente: null,
      clienteSearchText: "",
      isLoadingClientes: false,
      isLoadingAsuntos: false,
      isLoadingSubasuntos: false,
      isLoadingSubtipos: false,
      asuntos: [],
      selectedAsunto: null,
      subasuntos: [],
      selectedSubasunto: null,
      tiposDocumento: [],
      selectedTipoDocumento: null,
      subTiposDocumento: [],
      selectedSubTipoDocumento: null,
      carpetaRRHHOptions: [],
      selectedCarpetaRRHH: null,
      nivel1Options: [],
      selectedNivel1: null,
      nivel2Options: [],
      selectedNivel2: null,
      temasContaduria: [],
      selectedTema: null,
      subtemasContaduria: [],
      selectedSubtema: null,
      tiposDocumentoContaduria: [],
      selectedTipoDocumentoContaduria: null,
      carpetaDJOptions: [],
      selectedCarpetaDJ: null,
      carpeta1Options: [],
      selectedCarpeta1: null,
      carpeta2Options: [],
      selectedCarpeta2: null,
      carpeta3Options: [],
      selectedCarpeta3: null,
      carpeta4Options: [],
      selectedCarpeta4: null,
      carpeta5Options: [],
      selectedCarpeta5: null,
      carpeta6Options: [],
      selectedCarpeta6: null,
      carpeta7Options: [],
      selectedCarpeta7: null,
      isLoadingData: false,
      hasLoadedExistingMetadata: false,
    };
  }

  async componentDidMount() {
    if (!this.props.preservedState) {
      await this.loadDataForLibrary();
    } else {
      this.restoreState(this.props.preservedState);
    }
  }

  async componentDidUpdate(prevProps: MetadataComponentsProps) {
    if (prevProps.bibliotecaId !== this.props.bibliotecaId) {
      this.resetAllSelections();
      await this.loadDataForLibrary();
    } else if (this.props.preservedState && prevProps.preservedState !== this.props.preservedState) {
      this.restoreState(this.props.preservedState);
    } else if (this.props.existingMetadata && prevProps.existingMetadata !== this.props.existingMetadata && !this.state.hasLoadedExistingMetadata) {
      setTimeout(async () => {
        await this.loadExistingMetadata();
      }, 100);
    }
  }

  public restoreState = (preservedState: MetadataState) => {
    this.setState(
      {
        ...preservedState,
        isLoadingData: false,
        isLoadingClientes: false,
        isLoadingAsuntos: false,
        isLoadingSubasuntos: false,
        isLoadingSubtipos: false,
      },
      async () => {
        if (preservedState.selectedCliente && this.props.bibliotecaId.toUpperCase() === "DOCUMENTOS_CLIENTES") {
          await this.loadClientesDependentDataFixed(this.props.existingMetadata || {});
        }
        this.updateMetadata();
      }
    );
  };

  private resetAllSelections = () => {
    this.setState({
      selectedCliente: null,
      selectedAsunto: null,
      selectedSubasunto: null,
      selectedTipoDocumento: null,
      selectedSubTipoDocumento: null,
      selectedCarpetaRRHH: null,
      selectedNivel1: null,
      selectedNivel2: null,
      selectedTema: null,
      selectedSubtema: null,
      selectedTipoDocumentoContaduria: null,
      selectedCarpetaDJ: null,
      selectedCarpeta1: null,
      selectedCarpeta2: null,
      selectedCarpeta3: null,
      selectedCarpeta4: null,
      selectedCarpeta5: null,
      selectedCarpeta6: null,
      selectedCarpeta7: null,
      asuntos: [],
      subasuntos: [],
      subTiposDocumento: [],
      subtemasContaduria: [],
      carpeta2Options: [],
      carpeta3Options: [],
      carpeta4Options: [],
      carpeta5Options: [],
      carpeta6Options: [],
      carpeta7Options: [],
      isLoadingAsuntos: false,
      isLoadingSubasuntos: false,
      isLoadingSubtipos: false,
      clienteSearchText: "",
      hasLoadedExistingMetadata: false,
    });

  };

  private updateMetadata = () => {
    if (this.state.isLoadingData) return;
    const metadata = this.generateMetadata();
    this.props.onMetadataChange(metadata);
  };

  private generateMetadata = (): any => {
    const { bibliotecaId } = this.props;
    const metadata: any = {};

    switch (bibliotecaId.toUpperCase()) {
      case "DOCUMENTOS_CLIENTES":
        metadata.Cliente = this.state.selectedCliente?.title || "SIN CLASIFICAR";
        metadata.Asunto = this.state.selectedAsunto?.title || "SIN CLASIFICAR";
        metadata.S_Asunto = this.state.selectedSubasunto?.title || "SIN CLASIFICAR";
        metadata.Tipo_Doc = this.state.selectedTipoDocumento?.id || "";
        metadata.S_Tipo = this.state.selectedSubTipoDocumento?.title || "";
        break;

      case "DOCUMENTOS_ADMIN_RRHH":
        metadata.Carpeta1 = this.state.selectedCarpetaRRHH?.title || "SIN CLASIFICAR";
        break;

      case "DOCUMENTOS_CONSULADO_AUSTRALIA":
        const nivel1Value = this.state.selectedNivel1?.title || "SIN CLASIFICAR";
        const nivel2Value = this.state.selectedNivel2?.title || "SIN CLASIFICAR";
        metadata.Nivel1 = nivel1Value;
        metadata.Nivel2 = nivel2Value;
        break;

      case "DOCUMENTOS_CONTADURIA":
        const temaValue = this.state.selectedTema?.title || "SIN CLASIFICAR";
        const subtemaValue = this.state.selectedSubtema?.title || "SIN CLASIFICAR";
        const tipoDocSelected = this.state.selectedTipoDocumentoContaduria;

        let tipoDocValue = "";
        if (tipoDocSelected) {
          tipoDocValue = tipoDocSelected.id || tipoDocSelected.title || "SIN CLASIFICAR";
        }

        metadata.Tema = temaValue;
        metadata.SubTema = subtemaValue;
        metadata.TipoDoc = tipoDocValue;
        break;

      case "DOCUMENTOS_DECLARACIONES_JURADAS":
        const carpetaDJValue = this.state.selectedCarpetaDJ?.title || "SIN CLASIFICAR";
        metadata.Carpeta1 = carpetaDJValue;
        break;

      case "DOCUMENTOS_INTERNO":
        metadata.Carpeta1 = this.state.selectedCarpeta1?.title || "SIN CLASIFICAR";
        metadata.Carpeta2 = this.state.selectedCarpeta2?.title || "SIN CLASIFICAR";
        metadata.Carpeta3 = this.state.selectedCarpeta3?.title || "SIN CLASIFICAR";
        metadata.Carpeta4 = this.state.selectedCarpeta4?.title || "SIN CLASIFICAR";
        metadata.Carpeta5 = this.state.selectedCarpeta5?.title || "SIN CLASIFICAR";
        metadata.Carpeta6 = this.state.selectedCarpeta6?.title || "SIN CLASIFICAR";
        metadata.Carpeta7 = this.state.selectedCarpeta7?.title || "SIN CLASIFICAR";
        break;

      case "DOCUMENTOS_SOCIOS":
        break;
    }
    return metadata;
  };

  private handleClienteChange = async (_event: React.FormEvent<IComboBox>, option?: IComboBoxOption, _index?: number, value?: string) => {
    if (option) {
      // const cliente = { id: option.key as string, title: option.text };
      const realId = (option.key as string).split('-')[1];
      const cliente = { id: realId, title: option.text };

      this.setState({
        selectedCliente: cliente,
        clienteSearchText: cliente.title,
        asuntos: [],
        selectedAsunto: null,
        subasuntos: [],
        selectedSubasunto: null,
        isLoadingAsuntos: true,
      });

      try {
        const asuntos = await this.props.dataService.getAsuntosByCliente(cliente.id, cliente.title);
        this.setState({ asuntos, isLoadingAsuntos: false });
      } catch (error) {
        console.error("Error loading asuntos:", error);
        this.setState({ isLoadingAsuntos: false, asuntos: [] });
      }

      this.updateMetadata();

    } else if (value !== undefined) {
      this.setState({ clienteSearchText: value });
      this.filterClientes(value);
    }
  };


  private getAsuntoOptions = (): IDropdownOption[] => {
    return this.state.asuntos.map((a, index) => ({
      key: `asunto-${a.id}-${index}`,
      text: a.title,
      selected: this.state.selectedAsunto?.id === a.id
    }));
  };

  private getSubasuntoOptions = (): IDropdownOption[] => {
    return this.state.subasuntos.map(s => ({
      key: s.id,
      text: s.title,
      selected: this.state.selectedSubasunto?.id === s.id
    }));
  };

  private getSubtipoOptions = (): IDropdownOption[] => {
    return this.state.subTiposDocumento.map(s => ({
      key: s.id,
      text: s.title,
      selected: this.state.selectedSubTipoDocumento?.id === s.id
    }));
  };

  private handleClienteInputChange = (value: string) => {
    this.setState({ clienteSearchText: value });
    this.filterClientes(value);
  };

  //   private renderPerformanceInfo = () => {
  //   const cacheMetadata = this.props.dataService.getCacheMetadata();

  //   if (!cacheMetadata || this.props.bibliotecaId !== "DOCUMENTOS_CLIENTES") {
  //     return null;
  //   }

  //   return (
  //     <div style={{ 
  //       fontSize: '11px', 
  //       color: '#666', 
  //       marginBottom: '8px',
  //       padding: '4px 8px',
  //       backgroundColor: '#f0f8ff',
  //       borderRadius: '3px',
  //       border: '1px solid #e1e8ed'
  //     }}>
  //       ⚡ Datos cargados desde cache local - Rendimiento optimizado
  //       {cacheMetadata.isStale && (
  //         <span style={{ color: '#ff8c00' }}> • Actualizando en segundo plano</span>
  //       )}
  //     </div>
  //   );
  // };

  private getSelectedAsuntoKey = (): string | undefined => {
    if (!this.state.selectedAsunto) return undefined;

    const index = this.state.asuntos.findIndex(a => a.id === this.state.selectedAsunto?.id);
    return index >= 0 ? `asunto-${this.state.selectedAsunto.id}-${index}` : undefined;
  };

  private renderClientesFields = () => {
    const {
      clientesFiltered,
      selectedCliente,
      clienteSearchText,
      isLoadingClientes,
      isLoadingAsuntos,
      isLoadingSubasuntos,
      isLoadingSubtipos,
      asuntos,
      subasuntos,
      subTiposDocumento,
    } = this.state;
    // const clienteOptions: IComboBoxOption[] = clientesFiltered.map(c => ({ key: c.id, text: c.title }));
    const clienteOptions: IComboBoxOption[] = clientesFiltered.map((c, index) => ({
      key: `cliente-${c.id}-${index}`, // ← Agregar prefijo e índice
      text: c.title
    }));

    return (
      <>
        {/* {this.renderPerformanceInfo()} */}
        <div className="ms-u-marginBottom-16">
          {isLoadingClientes ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Spinner size={SpinnerSize.small} />
              <span>Cargando clientes...</span>
            </div>
          ) : (
            <ComboBox
              label="Cliente"
              placeholder="Buscar cliente..."
              allowFreeform={true}
              autoComplete="on"
              options={clienteOptions}
              selectedKey={selectedCliente?.id}
              text={clienteSearchText}
              onChange={this.handleClienteChange}
              onInputValueChange={this.handleClienteInputChange}
              onMenuOpen={() => {
                if (clientesFiltered.length === 0 && !clienteSearchText) this.filterClientes("");
              }}
              calloutProps={{
                calloutMaxHeight: 300,
                calloutMaxWidth: 400,
                isBeakVisible: false,
                directionalHint: 6,
              }}
              useComboBoxAsMenuWidth={false}
              required
            />
          )}

          {clientesFiltered.length === this.MAX_VISIBLE_ITEMS && (
            <div className="ms-fontSize-12 ms-fontColor-neutralSecondary ms-u-marginTop-4">
              Mostrando primeros {this.MAX_VISIBLE_ITEMS} resultados. Use la búsqueda para filtrar.
            </div>
          )}
        </div>

        {/* ASUNTO */}
        <div className="ms-u-marginBottom-16">
          {isLoadingAsuntos ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
              <Spinner size={SpinnerSize.small} />
              <span>Cargando asuntos para {selectedCliente?.title}...</span>
            </div>
          ) : (
            (asuntos.length > 0 || this.state.selectedAsunto) && (
              <Dropdown
                label="Asunto"
                placeholder="Seleccione un asunto..."
                options={this.getAsuntoOptions()}
                // selectedKey={this.state.selectedAsunto?.id}
                selectedKey={this.getSelectedAsuntoKey()}
                onChange={this.handleAsuntoSelect}
                calloutProps={{
                  calloutMaxHeight: 300,
                  calloutMaxWidth: 400,
                  isBeakVisible: false,
                  directionalHint: 6,
                }}
              />
            )
          )}
        </div>

        {/* SUB-ASUNTO */}
        <div className="ms-u-marginBottom-16">
          {isLoadingSubasuntos ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
              <Spinner size={SpinnerSize.small} />
              <span>Cargando sub-asuntos...</span>
            </div>
          ) : (
            (subasuntos.length > 0 || this.state.selectedSubasunto) && (
              <Dropdown
                label="Sub-Asunto"
                placeholder="Seleccione un sub-asunto..."
                options={this.getSubasuntoOptions()}
                selectedKey={this.state.selectedSubasunto?.id}
                onChange={this.handleSubasuntoSelect}
                calloutProps={{
                  calloutMaxHeight: 300,
                  calloutMaxWidth: 400,
                  isBeakVisible: false,
                  directionalHint: 6,
                }}
              />
            )
          )}
        </div>

        <Dropdown
          label="Tipo de Documento"
          placeholder="Seleccione tipo de documento..."
          options={this.state.tiposDocumento.map(t => ({ key: t.id, text: t.title }))}
          selectedKey={this.state.selectedTipoDocumento?.id}
          onChange={this.handleTipoDocumentoSelect}
          className="ms-u-marginBottom-16"
          calloutProps={{
            calloutMaxHeight: 300,
            calloutMaxWidth: 400,
            isBeakVisible: false,
            directionalHint: 6,
          }}
        />

        <div className="ms-u-marginBottom-16">
          {isLoadingSubtipos ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
              <Spinner size={SpinnerSize.small} />
              <span>Cargando sub-tipos...</span>
            </div>
          ) : (
            (subTiposDocumento.length > 0 || this.state.selectedSubTipoDocumento) && (
              <Dropdown
                label="Sub-Tipo de Documento"
                placeholder="Seleccione sub-tipo..."
                options={this.getSubtipoOptions()}
                selectedKey={this.state.selectedSubTipoDocumento?.id}
                onChange={this.handleSubTipoDocumentoSelect}
                // onChange={(_event, option) => {
                //   this.setState({ selectedSubTipoDocumento: option ? { id: option.key as string, title: option.text } : null });
                //   this.updateMetadata();
                // }}
                calloutProps={{
                  calloutMaxHeight: 300,
                  calloutMaxWidth: 400,
                  isBeakVisible: false,
                  directionalHint: 6,
                }}
              />
            )
          )}
        </div>
      </>
    );
  };


  private renderRRHHFields = () => (
    <Dropdown
      label="Carpeta RRHH"
      placeholder="Seleccione una carpeta..."
      options={this.state.carpetaRRHHOptions.map(c => ({ key: c.id, text: c.title }))}
      selectedKey={this.state.selectedCarpetaRRHH?.id}
      onChange={(_event, option) => {
        const newSelection = option ? { id: option.key as string, title: option.text } : null;
        this.setState({ selectedCarpetaRRHH: newSelection }, () => {
          this.updateMetadata();
        });
      }}
      className="ms-u-marginBottom-16"
      calloutProps={{
        calloutMaxHeight: 300,
        calloutMaxWidth: 400,
        isBeakVisible: false,
        directionalHint: 6,
      }}
    />
  );

  private renderConsuladoFields = () => {
    return (
      <>
        <Dropdown
          label="Nivel 1"
          placeholder="Seleccione nivel 1..."
          options={this.state.nivel1Options.map(n => ({ key: n.id, text: n.title }))}

          selectedKey={this.state.selectedNivel1?.id}

          onChange={(_event, option) => {

            const newSelection = option ? { id: option.key as string, title: option.text } : null;

            this.setState({ selectedNivel1: newSelection }, () => {

              this.updateMetadata();

            });

          }}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />



        <Dropdown

          label="Nivel 2"

          placeholder="Seleccione nivel 2..."

          options={this.state.nivel2Options.map(n => ({ key: n.id, text: n.title }))}

          selectedKey={this.state.selectedNivel2?.id}

          onChange={(_event, option) => {

            const newSelection = option ? { id: option.key as string, title: option.text } : null;

            this.setState({ selectedNivel2: newSelection }, () => {

              this.updateMetadata();

            });

          }}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      </>

    );

  };

  private renderContaduriaFields = () => {

    return (

      <>

        <Dropdown

          label="Tema"

          placeholder="Seleccione un tema..."

          options={this.state.temasContaduria.map(t => ({ key: t.id, text: t.title }))}

          selectedKey={this.state.selectedTema?.id}

          onChange={this.handleTemaSelect}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />



        {this.state.subtemasContaduria.length > 0 && (

          <Dropdown

            label="Sub-Tema"

            placeholder="Seleccione un sub-tema..."

            options={this.state.subtemasContaduria.map(s => ({ key: s.id, text: s.title }))}

            selectedKey={this.state.selectedSubtema?.id}

            onChange={(_event, option) => {

              const newSelection = option ? { id: option.key as string, title: option.text } : null;

              this.setState({ selectedSubtema: newSelection }, () => {

                this.updateMetadata();

              });

            }}

            className="ms-u-marginBottom-16"

            calloutProps={{

              calloutMaxHeight: 300,

              calloutMaxWidth: 400,

              isBeakVisible: false,

              directionalHint: 6,

            }}

          />

        )}



        <Dropdown

          label="Tipo de Documento"

          placeholder="Seleccione tipo de documento..."

          options={this.state.tiposDocumentoContaduria.map(t => ({ key: t.id, text: t.title }))}

          selectedKey={this.state.selectedTipoDocumentoContaduria?.id}

          onChange={(_event, option) => {

            if (option) {

              const newSelection = { id: option.key as string, title: option.text };

              this.setState({ selectedTipoDocumentoContaduria: newSelection }, () => {

                this.updateMetadata();

              });

            } else {

              this.setState({ selectedTipoDocumentoContaduria: null }, () => {

                this.updateMetadata();

              });

            }

          }}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      </>

    );

  };



  private renderDJFields = () => {

    return (

      <div>

        <Dropdown

          label="Carpeta DJ"

          placeholder="Seleccione una carpeta..."

          options={this.state.carpetaDJOptions.map(c => ({ key: c.id, text: c.title }))}

          selectedKey={this.state.selectedCarpetaDJ?.id}

          onChange={(_event, option) => {

            if (option) {

              const newSelection = { id: option.key as string, title: option.text };

              this.setState({ selectedCarpetaDJ: newSelection }, () => {

                this.updateMetadata();

              });

            } else {

              this.setState({ selectedCarpetaDJ: null }, () => {

                this.updateMetadata();

              });

            }

          }}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      </div>

    );

  };



  private renderInternoFields = () => (

    <>

      <Dropdown

        label="Carpeta 1"

        placeholder="Seleccione carpeta 1..."

        options={this.state.carpeta1Options.map(c => ({ key: c.id, text: c.title }))}

        selectedKey={this.state.selectedCarpeta1?.id}

        onChange={this.handleCarpeta1Select}

        className="ms-u-marginBottom-16"

        calloutProps={{

          calloutMaxHeight: 300,

          calloutMaxWidth: 400,

          isBeakVisible: false,

          directionalHint: 6,

        }}

      />

      {this.state.carpeta2Options.length > 0 && (

        <Dropdown

          label="Carpeta 2"

          placeholder="Seleccione carpeta 2..."

          options={this.state.carpeta2Options.map(c => ({ key: c.id, text: c.title }))}

          selectedKey={this.state.selectedCarpeta2?.id}

          onChange={this.handleCarpeta2Select}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      )}

      {this.state.carpeta3Options.length > 0 && (

        <Dropdown

          label="Carpeta 3"

          placeholder="Seleccione carpeta 3..."

          options={this.state.carpeta3Options.map(c => ({ key: c.id, text: c.title }))}

          selectedKey={this.state.selectedCarpeta3?.id}

          onChange={this.handleCarpeta3Select}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      )}

      {this.state.carpeta4Options.length > 0 && (

        <Dropdown

          label="Carpeta 4"

          placeholder="Seleccione carpeta 4..."

          options={this.state.carpeta4Options.map(c => ({ key: c.id, text: c.title }))}

          selectedKey={this.state.selectedCarpeta4?.id}

          onChange={this.handleCarpeta4Select}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      )}

      {this.state.carpeta5Options.length > 0 && (

        <Dropdown

          label="Carpeta 5"

          placeholder="Seleccione carpeta 5..."

          options={this.state.carpeta5Options.map(c => ({ key: c.id, text: c.title }))}

          selectedKey={this.state.selectedCarpeta5?.id}

          onChange={this.handleCarpeta5Select}

          className="ms-u-marginBottom-16"

          calloutProps={{

            calloutMaxHeight: 300,

            calloutMaxWidth: 400,

            isBeakVisible: false,

            directionalHint: 6,

          }}

        />

      )}

      {this.state.carpeta6Options.length > 0 && (
        <Dropdown
          label="Carpeta 6"
          placeholder="Seleccione carpeta 6..."
          options={this.state.carpeta6Options.map(c => ({ key: c.id, text: c.title }))}
          selectedKey={this.state.selectedCarpeta6?.id}
          onChange={this.handleCarpeta6Select}
          className="ms-u-marginBottom-16"
          calloutProps={{
            calloutMaxHeight: 300,
            calloutMaxWidth: 400,
            isBeakVisible: false,
            directionalHint: 6,
          }}
        />
      )}

      {this.state.carpeta7Options.length > 0 && (
        <Dropdown
          label="Carpeta 7"
          placeholder="Seleccione carpeta 7..."
          options={this.state.carpeta7Options.map(c => ({ key: c.id, text: c.title }))}
          selectedKey={this.state.selectedCarpeta7?.id}
          onChange={(_event, option) => {
            this.setState({ selectedCarpeta7: option ? { id: option.key as string, title: option.text } : null });
            this.updateMetadata();
          }}

          className="ms-u-marginBottom-16"
          calloutProps={{
            calloutMaxHeight: 300,
            calloutMaxWidth: 400,
            isBeakVisible: false,
            directionalHint: 6,
          }}
        />
      )}
    </>
  );

  private loadDataForLibrary = async () => {
    if (!this.props.bibliotecaId) {
      return;
    }

    this.setState({ isLoadingData: true });
    try {
      switch (this.props.bibliotecaId) {
        case "DOCUMENTOS_CLIENTES":
          await this.loadClientesData();
          break;

        case "DOCUMENTOS_SOCIOS":
          break;

        case "DOCUMENTOS_ADMIN_RRHH":
          const carpetaRRHH = await this.props.dataService.getCarpetaRRHH();
          this.setState({ carpetaRRHHOptions: carpetaRRHH });
          break;

        case "DOCUMENTOS_CONSULADO_AUSTRALIA":
          const [nivel1Options, nivel2Options] = await Promise.all([
            this.props.dataService.getNivel1(),
            this.props.dataService.getNivel2(),
          ]);
          this.setState({ nivel1Options, nivel2Options });
          break;

        case "DOCUMENTOS_CONTADURIA":
          const [temasContaduria, tiposDocumentoContaduria] = await Promise.all([
            this.props.dataService.getTemasContaduria(),
            this.props.dataService.getTipoDocContaduria(),
          ]);
          this.setState({ temasContaduria, tiposDocumentoContaduria });
          break;

        case "DOCUMENTOS_DECLARACIONES_JURADAS":
          const carpetaDJ = await this.props.dataService.getCarpetaDJ();
          this.setState({ carpetaDJOptions: carpetaDJ });
          break;

        case "DOCUMENTOS_INTERNO":
          const carpeta1 = await this.props.dataService.getCarpeta1();
          this.setState({ carpeta1Options: carpeta1 });
          break;
      }
      this.setState({ isLoadingData: false }, () => {
        if (this.props.existingMetadata && !this.state.hasLoadedExistingMetadata) {
          setTimeout(async () => {
            await this.loadExistingMetadata();
          }, 100);
        }
      });
    } catch (error) {
      console.error("Error loading library data:", error);
      this.setState({ isLoadingData: false });
    }
  };

  private loadClientesData = async () => {
    this.setState({ isLoadingClientes: true });
    try {
      const [clientes, tiposDocumento] = await Promise.all([
        this.props.dataService.getClientes("", 6000),
        this.props.dataService.getTiposDocumento(),
      ]);

      const clientesOrdenados = clientes.sort((a, b) => a.title.localeCompare(b.title));
      this.setState({
        clientes: clientesOrdenados,
        clientesFiltered: clientesOrdenados.slice(0, this.MAX_VISIBLE_ITEMS),
        tiposDocumento,
      });
    } catch (error) {
      console.error("Error loading clientes:", error);
    } finally {
      this.setState({ isLoadingClientes: false });
    }
  };

  private loadExistingMetadata = async () => {
    const { existingMetadata, bibliotecaId } = this.props;
    if (!existingMetadata || this.state.hasLoadedExistingMetadata) return;
    const newState: Partial<MetadataState> = {};
    try {
      switch (bibliotecaId.toUpperCase()) {
        case "DOCUMENTOS_CLIENTES":
          await this.loadExistingClientesMetadata(existingMetadata, newState);
          break;

        case "DOCUMENTOS_ADMIN_RRHH":
          await this.loadExistingRRHHMetadata(existingMetadata, newState);
          break;

        case "DOCUMENTOS_CONSULADO_AUSTRALIA":
          await this.loadExistingConsuladoMetadata(existingMetadata, newState);
          break;

        case "DOCUMENTOS_CONTADURIA":
          await this.loadExistingContaduriaMetadata(existingMetadata, newState);
          break;

        case "DOCUMENTOS_DECLARACIONES_JURADAS":
          await this.loadExistingDJMetadata(existingMetadata, newState);
          break;

        case "DOCUMENTOS_INTERNO":
          await this.loadExistingInternoMetadata(existingMetadata, newState);
          break;

        case "DOCUMENTOS_SOCIOS":
          break;

        default:
          console.warn("Biblioteca no reconocida:", bibliotecaId);
          break;
      }

      this.setState({
        ...this.state,
        ...newState,
        hasLoadedExistingMetadata: true,
      }, async () => {
        if (bibliotecaId.toUpperCase() === "DOCUMENTOS_CLIENTES") {
          this.loadClientesDependentDataFixed(existingMetadata);
        } else if (bibliotecaId.toUpperCase() === "DOCUMENTOS_INTERNO") {
          await this.loadInternoDependentData(existingMetadata);
        }

        setTimeout(() => {
          this.updateMetadata();
        }, 100);
      });

    } catch (error) {
      console.error("Error cargando metadatos existentes:", error);
      this.setState({ hasLoadedExistingMetadata: true });
    }
  };

  // MÉTODO HELPER: Buscar ID real en una lista de opciones por título
  private findIdByTitle = (options: any[], title: string): string | null => {
    if (!options || !title) return null;

    const found = options.find(option =>
      option.title?.toLowerCase().trim() === title.toLowerCase().trim() ||
      option.id?.toLowerCase().trim() === title.toLowerCase().trim()
    );

    return found ? found.id : null;
  };

  private loadInternoDependentData = async (existingMetadata: any) => {
    try {
      console.log("Cargando carpetas dependientes para INTERNO...");
      console.log("ExistingMetadata:", existingMetadata);
      console.log("Estado actual de carpetas:", {
        carpeta1: this.state.selectedCarpeta1,
        carpeta2: this.state.selectedCarpeta2,
        carpeta3: this.state.selectedCarpeta3,
        carpeta4: this.state.selectedCarpeta4,
        carpeta5: this.state.selectedCarpeta5,
        carpeta6: this.state.selectedCarpeta6,
        carpeta7: this.state.selectedCarpeta7
      });

      if (this.state.selectedCarpeta1) {
        try {
          const carpeta2Options = await this.props.dataService.getCarpeta2ByCarpeta1(this.state.selectedCarpeta1.id);
          await new Promise<void>((resolve) => {
            this.setState({ carpeta2Options }, () => {
              console.log("Carpeta2 opciones establecidas en el estado");
              resolve();
            });
          });

          if (this.state.selectedCarpeta2 && existingMetadata.Carpeta2) {
            const carpeta2RealId = this.findIdByTitle(carpeta2Options, existingMetadata.Carpeta2);

            if (carpeta2RealId) {
              await new Promise<void>((resolve) => {
                this.setState({
                  selectedCarpeta2: {
                    id: carpeta2RealId,
                    title: existingMetadata.Carpeta2
                  }
                }, () => {
                  resolve();
                });
              });

              await this.loadCarpeta3AndBeyond(existingMetadata);
            }
          }
        } catch (error) {
          console.error("Error loading carpeta2 options:", error);
        }
      }
    } catch (error) {
      console.error("Error en loadInternoDependentData:", error);
    }
  };

  private loadCarpeta3AndBeyond = async (existingMetadata: any) => {
    try {
      // Cargar Carpeta3
      if (this.state.selectedCarpeta2) {
        const carpeta3Options = await this.props.dataService.getCarpeta3ByCarpeta2(this.state.selectedCarpeta2.id, this.state.selectedCarpeta2.title);

        await new Promise<void>((resolve) => {
          this.setState({ carpeta3Options }, () => {
            resolve();
          });
        });

        // Buscar ID real de Carpeta3
        if (this.state.selectedCarpeta3 && existingMetadata.Carpeta3) {
          const carpeta3RealId = this.findIdByTitle(carpeta3Options, existingMetadata.Carpeta3);

          if (carpeta3RealId) {
            await new Promise<void>((resolve) => {
              this.setState({
                selectedCarpeta3: {
                  id: carpeta3RealId,
                  title: existingMetadata.Carpeta3
                }
              }, () => {
                resolve();
              });
            });

            // Cargar Carpeta4
            const carpeta4Options = await this.props.dataService.getCarpeta4ByCarpeta3(carpeta3RealId, existingMetadata.Carpeta3);

            await new Promise<void>((resolve) => {
              this.setState({ carpeta4Options }, () => {
                resolve();
              });
            });

            // Buscar ID real de Carpeta4
            if (this.state.selectedCarpeta4 && existingMetadata.Carpeta4) {
              const carpeta4RealId = this.findIdByTitle(carpeta4Options, existingMetadata.Carpeta4);

              if (carpeta4RealId) {
                await new Promise<void>((resolve) => {
                  this.setState({
                    selectedCarpeta4: {
                      id: carpeta4RealId,
                      title: existingMetadata.Carpeta4
                    }
                  }, () => {
                    resolve();
                  });
                });

                // Cargar Carpeta5
                const carpeta5Options = await this.props.dataService.getCarpeta5ByCarpeta4(carpeta4RealId, existingMetadata.Carpeta4);

                await new Promise<void>((resolve) => {
                  this.setState({ carpeta5Options }, () => {
                    resolve();
                  });
                });

                // Buscar ID real de Carpeta5
                if (this.state.selectedCarpeta5 && existingMetadata.Carpeta5) {
                  const carpeta5RealId = this.findIdByTitle(carpeta5Options, existingMetadata.Carpeta5);

                  if (carpeta5RealId) {
                    await new Promise<void>((resolve) => {
                      this.setState({
                        selectedCarpeta5: {
                          id: carpeta5RealId,
                          title: existingMetadata.Carpeta5
                        }
                      }, () => {
                        resolve();
                      });
                    });

                    // Cargar Carpeta6
                    const carpeta6Options = await this.props.dataService.getCarpeta6ByCarpeta5(carpeta5RealId);

                    await new Promise<void>((resolve) => {
                      this.setState({ carpeta6Options }, () => {
                        resolve();
                      });
                    });

                    // Buscar ID real de Carpeta6
                    if (this.state.selectedCarpeta6 && existingMetadata.Carpeta6) {
                      const carpeta6RealId = this.findIdByTitle(carpeta6Options, existingMetadata.Carpeta6);

                      if (carpeta6RealId) {
                        await new Promise<void>((resolve) => {
                          this.setState({
                            selectedCarpeta6: {
                              id: carpeta6RealId,
                              title: existingMetadata.Carpeta6
                            }
                          }, () => {
                            resolve();
                          });
                        });

                        // Cargar Carpeta7
                        const carpeta7Options = await this.props.dataService.getCarpeta7ByCarpeta6(carpeta6RealId);

                        await new Promise<void>((resolve) => {
                          this.setState({ carpeta7Options }, () => {
                            resolve();
                          });
                        });
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    } catch (error) {
      console.error("Error en loadCarpeta3AndBeyond:", error);
    }
  };

  private loadExistingClientesMetadata = async (existingMetadata: any, newState: Partial<MetadataState>) => {
    // Cliente
    if (existingMetadata.Cliente) {
      const clienteEnLista = this.state.clientes.find(c => {
        const match = c.title.toLowerCase().trim() === existingMetadata.Cliente.toLowerCase().trim();
        return match;
      });

      if (clienteEnLista) {
        newState.selectedCliente = {
          id: clienteEnLista.id,
          title: clienteEnLista.title
        };
        newState.clienteSearchText = clienteEnLista.title;
      } else {
        newState.selectedCliente = {
          id: existingMetadata.Cliente,
          title: existingMetadata.Cliente
        };
        newState.clienteSearchText = existingMetadata.Cliente;
      }
    } else {
      console.log("No hay existingMetadata.Cliente");
    }

    // Asunto
    if (existingMetadata.Asunto) {
      newState.selectedAsunto = {
        id: existingMetadata.Asunto,
        title: existingMetadata.Asunto
      };
    } else {
      console.log("No hay existingMetadata.Asunto");
    }

    // Subasunto
    if (existingMetadata.S_Asunto) {
      newState.selectedSubasunto = {
        id: existingMetadata.S_Asunto,
        title: existingMetadata.S_Asunto
      };
    } else {
      console.log("No hay existingMetadata.S_Asunto");
    }

    // Tipo Documento
    if (existingMetadata.Tipo_Doc) {
      const tipoEnLista = this.state.tiposDocumento.find(t => {
        const match = t.id.toLowerCase().trim() === existingMetadata.Tipo_Doc.toLowerCase().trim();
        return match;
      });

      if (tipoEnLista) {
        newState.selectedTipoDocumento = tipoEnLista;
      } else {
        newState.selectedTipoDocumento = {
          id: existingMetadata.Tipo_Doc,
          title: existingMetadata.Tipo_Doc
        };
      }
    } else {
      console.log("No hay existingMetadata.Tipo_Doc");
    }

    // Subtipo Documento
    if (existingMetadata.S_Tipo) {
      newState.selectedSubTipoDocumento = {
        id: existingMetadata.S_Tipo,
        title: existingMetadata.S_Tipo
      };;
    } else {
      console.log("No hay existingMetadata.S_Tipo");
    }
  };

  private loadExistingRRHHMetadata = async (existingMetadata: any, newState: Partial<MetadataState>) => {
    if (existingMetadata.Carpeta1) {
      const carpetaEnLista = this.state.carpetaRRHHOptions.find(c =>
        c.title.toLowerCase().trim() === existingMetadata.Carpeta1.toLowerCase().trim()
      );
      if (carpetaEnLista) {
        newState.selectedCarpetaRRHH = {
          id: carpetaEnLista.id,
          title: carpetaEnLista.title
        };
      } else {
        newState.selectedCarpetaRRHH = {
          id: existingMetadata.Carpeta1,
          title: existingMetadata.Carpeta1
        };
      }
    }

    if (existingMetadata.CarpetaRRHH && !newState.selectedCarpetaRRHH) {
      const carpetaEnLista = this.state.carpetaRRHHOptions.find(c =>
        c.title.toLowerCase().trim() === existingMetadata.CarpetaRRHH.toLowerCase().trim() ||
        c.id.toLowerCase().trim() === existingMetadata.CarpetaRRHH.toLowerCase().trim()
      );

      if (carpetaEnLista) {
        newState.selectedCarpetaRRHH = {
          id: carpetaEnLista.id,
          title: carpetaEnLista.title
        };
      } else {
        newState.selectedCarpetaRRHH = {
          id: existingMetadata.CarpetaRRHH,
          title: existingMetadata.CarpetaRRHH
        };
      }
    }
  };

  private loadExistingConsuladoMetadata = async (existingMetadata: any, newState: Partial<MetadataState>) => {
    // Nivel 1
    if (existingMetadata.Nivel1) {
      const nivel1EnLista = this.state.nivel1Options.find(n =>
        n.title.toLowerCase().trim() === existingMetadata.Nivel1.toLowerCase().trim()
      );

      if (nivel1EnLista) {
        newState.selectedNivel1 = {
          id: nivel1EnLista.id,
          title: nivel1EnLista.title
        };
      } else {
        newState.selectedNivel1 = {
          id: existingMetadata.Nivel1,
          title: existingMetadata.Nivel1
        };
      }
    }

    // Nivel 2
    if (existingMetadata.Nivel2) {
      const nivel2EnLista = this.state.nivel2Options.find(n =>
        n.title.toLowerCase().trim() === existingMetadata.Nivel2.toLowerCase().trim()
      );

      if (nivel2EnLista) {
        newState.selectedNivel2 = {
          id: nivel2EnLista.id,
          title: nivel2EnLista.title
        };
      } else {
        newState.selectedNivel2 = {
          id: existingMetadata.Nivel2,
          title: existingMetadata.Nivel2
        };
      }
    }
  };

  private loadExistingContaduriaMetadata = async (existingMetadata: any, newState: Partial<MetadataState>) => {
    // Tema
    if (existingMetadata.Tema) {
      const temaEnLista = this.state.temasContaduria.find(t =>
        t.title.toLowerCase().trim() === existingMetadata.Tema.toLowerCase().trim()
      );

      if (temaEnLista) {
        newState.selectedTema = {
          id: temaEnLista.id,
          title: temaEnLista.title
        };
      } else {
        newState.selectedTema = {
          id: existingMetadata.Tema,
          title: existingMetadata.Tema
        };
      }
    }

    // SubTema
    if (existingMetadata.SubTema) {
      newState.selectedSubtema = {
        id: existingMetadata.SubTema,
        title: existingMetadata.SubTema
      };
    }

    // TipoDoc
    if (existingMetadata.TipoDoc) {
      const tipoDocEnLista = this.state.tiposDocumentoContaduria.find(t =>
        t.title.toLowerCase().trim() === existingMetadata.TipoDoc.toLowerCase().trim()
      );

      if (tipoDocEnLista) {
        newState.selectedTipoDocumentoContaduria = {
          id: tipoDocEnLista.id,
          title: tipoDocEnLista.title
        };
      } else {
        newState.selectedTipoDocumentoContaduria = {
          id: existingMetadata.TipoDoc,
          title: existingMetadata.TipoDoc
        };
      }
    }
  };

  private loadExistingDJMetadata = async (existingMetadata: any, newState: Partial<MetadataState>) => {
    if (existingMetadata.Carpeta1) {
      const carpetaEnLista = this.state.carpetaDJOptions.find(c =>
        c.title.toLowerCase().trim() === existingMetadata.Carpeta1.toLowerCase().trim()
      );

      if (carpetaEnLista) {
        newState.selectedCarpetaDJ = {
          id: carpetaEnLista.id,
          title: carpetaEnLista.title
        };
      } else {
        newState.selectedCarpetaDJ = {
          id: existingMetadata.Carpeta1,
          title: existingMetadata.Carpeta1
        };
      }
    }
  };

  private loadExistingInternoMetadata = async (existingMetadata: any, newState: Partial<MetadataState>) => {
    if (existingMetadata.Carpeta1) {
      const carpeta1EnLista = this.state.carpeta1Options.find(c =>
        c.title.toLowerCase().trim() === existingMetadata.Carpeta1.toLowerCase().trim()
      );

      if (carpeta1EnLista) {
        newState.selectedCarpeta1 = {
          id: carpeta1EnLista.id,
          title: carpeta1EnLista.title
        };
      } else {
        newState.selectedCarpeta1 = {
          id: existingMetadata.Carpeta1,
          title: existingMetadata.Carpeta1
        };
      }
    }

    if (existingMetadata.Carpeta2) {
      newState.selectedCarpeta2 = {
        id: existingMetadata.Carpeta2,
        title: existingMetadata.Carpeta2
      };
    }

    if (existingMetadata.Carpeta3) {
      newState.selectedCarpeta3 = {
        id: existingMetadata.Carpeta3,
        title: existingMetadata.Carpeta3
      };
    }

    if (existingMetadata.Carpeta4) {
      newState.selectedCarpeta4 = {
        id: existingMetadata.Carpeta4,
        title: existingMetadata.Carpeta4
      };
    }

    if (existingMetadata.Carpeta5) {
      newState.selectedCarpeta5 = {
        id: existingMetadata.Carpeta5,
        title: existingMetadata.Carpeta5
      };
    }

    if (existingMetadata.Carpeta6) {
      newState.selectedCarpeta6 = {
        id: existingMetadata.Carpeta6,
        title: existingMetadata.Carpeta6
      };
    }

    if (existingMetadata.Carpeta7) {
      newState.selectedCarpeta7 = {
        id: existingMetadata.Carpeta7,
        title: existingMetadata.Carpeta7
      };
    }

    console.log("Estado final para Interno:", {
      carpeta1: newState.selectedCarpeta1,
      carpeta2: newState.selectedCarpeta2,
      carpeta3: newState.selectedCarpeta3,
      carpeta4: newState.selectedCarpeta4,
      carpeta5: newState.selectedCarpeta5,
      carpeta6: newState.selectedCarpeta6,
      carpeta7: newState.selectedCarpeta7
    });
  };

  private loadClientesDependentDataFixed = async (existingMetadata: any) => {
    try {
      if (this.state.selectedCliente) {
        this.setState({ isLoadingAsuntos: true });
        const asuntos = await this.props.dataService.getAsuntosByCliente(
          this.state.selectedCliente.id,
          this.state.selectedCliente.title
        );

        let selectedAsunto = null;
        if (existingMetadata.Asunto) {
          selectedAsunto = asuntos.find(a =>
            a.title.toLowerCase().trim() === existingMetadata.Asunto.toLowerCase().trim() ||
            a.id.toLowerCase().trim() === existingMetadata.Asunto.toLowerCase().trim()
          );
        }

        this.setState({
          asuntos,
          selectedAsunto: selectedAsunto || this.state.selectedAsunto,
          isLoadingAsuntos: false
        }, async () => {
          const currentAsunto = selectedAsunto || this.state.selectedAsunto;
          if (currentAsunto && existingMetadata.S_Asunto) {
            this.setState({ isLoadingSubasuntos: true });
            const subasuntos = await this.props.dataService.getSubasuntosByAsuntoOriginal(
              currentAsunto.id,
              currentAsunto.title
            );

            let selectedSubasunto = null;
            if (existingMetadata.S_Asunto) {
              selectedSubasunto = subasuntos.find(s =>
                s.title.toLowerCase().trim() === existingMetadata.S_Asunto.toLowerCase().trim() ||
                s.id.toLowerCase().trim() === existingMetadata.S_Asunto.toLowerCase().trim()
              );
            }

            this.setState({
              subasuntos,
              selectedSubasunto: selectedSubasunto || this.state.selectedSubasunto,
              isLoadingSubasuntos: false
            });
          }
        });
      }

      const currentTipoDoc = this.state.selectedTipoDocumento;
      if (currentTipoDoc && existingMetadata.S_Tipo) {
        this.setState({ isLoadingSubtipos: true });
        const subTipos = await this.props.dataService.getSubTiposDocumento(
          currentTipoDoc.id
        );

        let selectedSubTipo = null;
        if (existingMetadata.S_Tipo) {
          selectedSubTipo = subTipos.find(s =>
            s.title.toLowerCase().trim() === existingMetadata.S_Tipo.toLowerCase().trim() ||
            s.id.toLowerCase().trim() === existingMetadata.S_Tipo.toLowerCase().trim()
          );
        }

        this.setState({
          subTiposDocumento: subTipos,
          selectedSubTipoDocumento: selectedSubTipo || this.state.selectedSubTipoDocumento,
          isLoadingSubtipos: false
        });
      }
    } catch (error) {
      console.error("Error cargando datos dependientes:", error);
      this.setState({
        isLoadingAsuntos: false,
        isLoadingSubasuntos: false,
        isLoadingSubtipos: false
      });
    }
  };

  private filterClientes = (searchText: string) => {
    if (this.searchTimeoutId) {
      clearTimeout(this.searchTimeoutId);
    }
    this.searchTimeoutId = setTimeout(() => {
      const { clientes } = this.state;
      const normalizedSearch = this.normalizeText(searchText);
      if (!normalizedSearch) {
        this.setState({ clientesFiltered: clientes.slice(0, this.MAX_VISIBLE_ITEMS) });
        return;
      }
      const filtered = clientes
        .filter(cliente =>
          this.normalizeText(cliente.title).includes(normalizedSearch) ||
          this.normalizeText(cliente.id).includes(normalizedSearch)
        )
        .slice(0, this.MAX_VISIBLE_ITEMS);
      this.setState({ clientesFiltered: filtered });
    }, this.SEARCH_DELAY);
  };

  private normalizeText = (text: string): string => {
    return text
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
  };

  public refreshData = async () => {
    // Resetear todas las selecciones
    this.resetAllSelections();
    
    // Recargar datos para la biblioteca actual
    await this.loadDataForLibrary();
    
    // Si hay metadatos existentes, recargarlos también
    if (this.props.existingMetadata) {
      this.setState({ hasLoadedExistingMetadata: false });
      setTimeout(async () => {
        await this.loadExistingMetadata();
      }, 100);
    }
  };

  private handleAsuntoSelect = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const realId = option ? (option.key as string).split('-')[1] : null;
    const asunto = option ? { id: realId!, title: option.text } : null;
    this.setState({
      selectedAsunto: asunto,
      subasuntos: [],
      selectedSubasunto: null,
      isLoadingSubasuntos: asunto ? true : false,
    });

    if (asunto) {
      try {
        const subasuntos = await this.props.dataService.getSubasuntosByAsunto(asunto.id, asunto.title);
        this.setState({
          subasuntos,
          isLoadingSubasuntos: false,
        });
      } catch (error) {
        this.setState({
          isLoadingSubasuntos: false,
          subasuntos: [],
        });
      }
    }

    this.updateMetadata();
  };

  private handleSubasuntoSelect = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const subasunto = option ? { id: option.key as string, title: option.text } : null;
    this.setState({ selectedSubasunto: subasunto });
    this.updateMetadata();
  };

  private handleTipoDocumentoSelect = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const tipoDocumento = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedTipoDocumento: tipoDocumento,
      subTiposDocumento: [],
      selectedSubTipoDocumento: null,
      isLoadingSubtipos: tipoDocumento ? true : false,
    });

    if (tipoDocumento) {
      try {
        const subTipos = await this.props.dataService.getSubTiposDocumento(tipoDocumento.id);

        this.setState({
          subTiposDocumento: subTipos,
          isLoadingSubtipos: false,
        });
      } catch (error) {
        console.error("Error loading subTipos:", error);
        this.setState({
          isLoadingSubtipos: false,
          subTiposDocumento: [],
        });
      }
    }

    this.updateMetadata();
  };

  private handleSubTipoDocumentoSelect = (_event, option) => {
    const newSubtipo = option ? { id: option.key as string, title: option.text } : null;
    this.setState({ selectedSubTipoDocumento: newSubtipo }, () => {
      this.updateMetadata();
    });
  };

  private handleTemaSelect = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const tema = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedTema: tema,
      subtemasContaduria: [],
      selectedSubtema: null,
    });

    if (tema) {
      try {
        const subtemas = await this.props.dataService.getSubtemasByTema(tema.id);
        this.setState({ subtemasContaduria: subtemas });
      } catch (error) {
        console.error("Error loading subtemas:", error);
      }
    }
    this.updateMetadata();
  };

  private handleCarpeta1Select = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const carpeta1 = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedCarpeta1: carpeta1,
      carpeta2Options: [],
      selectedCarpeta2: null,
      carpeta3Options: [],
      selectedCarpeta3: null,
      carpeta4Options: [],
      selectedCarpeta4: null,
      carpeta5Options: [],
      selectedCarpeta5: null,
      carpeta6Options: [],
      selectedCarpeta6: null,
      carpeta7Options: [],
      selectedCarpeta7: null,
    });

    if (carpeta1) {
      try {
        const carpeta2Options = await this.props.dataService.getCarpeta2ByCarpeta1(carpeta1.id);
        this.setState({ carpeta2Options });
      } catch (error) {
        console.error("Error loading carpeta2:", error);
      }
    }
    this.updateMetadata();
  };

  private handleCarpeta2Select = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const carpeta2 = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedCarpeta2: carpeta2,
      carpeta3Options: [],
      selectedCarpeta3: null,
      carpeta4Options: [],
      selectedCarpeta4: null,
      carpeta5Options: [],
      selectedCarpeta5: null,
      carpeta6Options: [],
      selectedCarpeta6: null,
      carpeta7Options: [],
      selectedCarpeta7: null,
    });

    if (carpeta2) {
      try {
        const carpeta3Options = await this.props.dataService.getCarpeta3ByCarpeta2(carpeta2.id, carpeta2.title);
        this.setState({ carpeta3Options });
      } catch (error) {
        console.error("Error loading carpeta3:", error);
      }
    }
    this.updateMetadata();
  };

  private handleCarpeta3Select = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const carpeta3 = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedCarpeta3: carpeta3,
      carpeta4Options: [],
      selectedCarpeta4: null,
      carpeta5Options: [],
      selectedCarpeta5: null,
      carpeta6Options: [],
      selectedCarpeta6: null,
      carpeta7Options: [],
      selectedCarpeta7: null,
    });

    if (carpeta3) {
      try {
        const carpeta4Options = await this.props.dataService.getCarpeta4ByCarpeta3(carpeta3.id, carpeta3.title);
        this.setState({ carpeta4Options });
      } catch (error) {
        console.error("Error loading carpeta4:", error);
      }
    }
    this.updateMetadata();
  };

  private handleCarpeta4Select = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const carpeta4 = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedCarpeta4: carpeta4,
      carpeta5Options: [],
      selectedCarpeta5: null,
      carpeta6Options: [],
      selectedCarpeta6: null,
      carpeta7Options: [],
      selectedCarpeta7: null,
    });

    if (carpeta4) {
      try {
        const carpeta5Options = await this.props.dataService.getCarpeta5ByCarpeta4(carpeta4.id, carpeta4.title);
        this.setState({ carpeta5Options });
      } catch (error) {
        console.error("Error loading carpeta5:", error);
      }
    }
    this.updateMetadata();
  };

  private handleCarpeta5Select = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const carpeta5 = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedCarpeta5: carpeta5,
      carpeta6Options: [],
      selectedCarpeta6: null,
      carpeta7Options: [],
      selectedCarpeta7: null,
    });

    if (carpeta5) {
      try {
        const carpeta6Options = await this.props.dataService.getCarpeta6ByCarpeta5(carpeta5.id);
        this.setState({ carpeta6Options });
      } catch (error) {
        console.error("Error loading carpeta6:", error);
      }
    }
    this.updateMetadata();
  };

  private handleCarpeta6Select = async (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const carpeta6 = option ? { id: option.key as string, title: option.text } : null;
    this.setState({
      selectedCarpeta6: carpeta6,
      carpeta7Options: [],
      selectedCarpeta7: null,
    });

    if (carpeta6) {
      try {
        const carpeta7Options = await this.props.dataService.getCarpeta7ByCarpeta6(carpeta6.id);
        this.setState({ carpeta7Options });
      } catch (error) {
        console.error("Error loading carpeta7:", error);
      }
    }
    this.updateMetadata();
  };

  render() {
    const { bibliotecaId, isLoading } = this.props;
    const { isLoadingData, isLoadingClientes } = this.state;
    const isAnyLoading = isLoading || isLoadingData || isLoadingClientes;

    if (!bibliotecaId || isAnyLoading) {
      return (
        <div style={{ textAlign: 'center', padding: '20px' }}>
          <Spinner size={SpinnerSize.medium} label="Cargando metadatos..." />
        </div>
      );
    }

    switch (bibliotecaId) {
      case "DOCUMENTOS_CLIENTES":
        return this.renderClientesFields();
      case "DOCUMENTOS_ADMIN_RRHH":
        return this.renderRRHHFields();
      case "DOCUMENTOS_CONSULADO_AUSTRALIA":
        return this.renderConsuladoFields();
      case "DOCUMENTOS_CONTADURIA":
        return this.renderContaduriaFields();
      case "DOCUMENTOS_DECLARACIONES_JURADAS":
        return this.renderDJFields();
      case "DOCUMENTOS_INTERNO":
        return this.renderInternoFields();
      default:
        return null;
    }
  }
}

const arePropsEqual = (prevProps: MetadataComponentsProps, nextProps: MetadataComponentsProps) => {
  const shouldRerender =
    prevProps.bibliotecaId !== nextProps.bibliotecaId ||
    prevProps.dataService !== nextProps.dataService ||
    prevProps.existingMetadata !== nextProps.existingMetadata ||
    prevProps.preservedState !== nextProps.preservedState ||
    (prevProps.isLoading !== nextProps.isLoading && nextProps.isLoading === false);

  return !shouldRerender;
};

export default React.memo(MetadataComponents, arePropsEqual);