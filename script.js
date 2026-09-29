document.addEventListener('DOMContentLoaded', () => {
    console.log('YanDEV - Plataforma carregada com sucesso.');

    // Inicialização do Lenis (Smooth Scroll)
    const lenis = new Lenis({
        duration: 1.4,
        easing: (t) => Math.min(1, 1.001 - Math.pow(2, -10 * t)),
        direction: 'vertical',
        gestureDirection: 'vertical',
        smooth: true,
    });

    function raf(time) {
        lenis.raf(time);
        requestAnimationFrame(raf);
    }
    requestAnimationFrame(raf);

    // Links âncora via Lenis
    const links = document.querySelectorAll('a[href^="#"]');
    links.forEach(link => {
        link.addEventListener('click', function(e) {
            const targetId = this.getAttribute('href');
            if (targetId && targetId.length > 1) {
                e.preventDefault();
                const targetElement = document.querySelector(targetId);
                if (targetElement) {
                    lenis.scrollTo(targetElement, { offset: 0, duration: 1.5 });
                }
            }
        });
    });

    // ============================================
    // --- FUNCIONALIDADE PPTX ---
    // ============================================
    const downloadPptxBtn = document.getElementById('downloadPptxBtn');
    const pptxBtnText = document.getElementById('pptxBtnText');

    if (downloadPptxBtn) {
        downloadPptxBtn.addEventListener('click', () => {
            downloadPptxBtn.style.opacity = '0.7';
            if (pptxBtnText) pptxBtnText.textContent = 'GERANDO...';

            setTimeout(async () => {
                try {
                    await generatePPTX();
                } catch (err) {
                    console.error('Erro ao gerar PPTX:', err);
                    alert('Erro ao gerar PPTX. Verifique o console.');
                } finally {
                    downloadPptxBtn.style.opacity = '1';
                    if (pptxBtnText) pptxBtnText.textContent = 'BAIXAR PPTX';
                }
            }, 300);
        });
    }

    // --- FUNÇÃO AUXILIAR: CONVERTER IMAGEM PARA BASE64 ---
    async function imageToBase64(url) {
        try {
            const response = await fetch(url);
            if (!response.ok) throw new Error('HTTP ' + response.status);
            const blob = await response.blob();
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        } catch (err) {
            console.warn('Não foi possível carregar imagem:', url, err);
            return null;
        }
    }

    async function generatePPTX() {
        const pptx = new PptxGenJS();

        pptx.layout = 'LAYOUT_16x9';
        pptx.author = 'Yan Rocha';
        pptx.company = 'YanDEV';
        pptx.title = 'Proposta Comercial - YanDEV';
        pptx.subject = 'Proposta Comercial';

        const COLORS = {
            black: '0B0B0B',
            white: 'FFFFFF',
            yellow: 'D4FF00',
            green: 'BDF57B',
            darkGray: '1A1A1A',
            midGray: '666666',
            lightGray: 'F7F6F2',
            navy: '0B1120',
            blue: '5F72FF'
        };

        // ==================================================
        // PRÉ-CARREGAR IMAGENS
        // ==================================================
        if (pptxBtnText) pptxBtnText.textContent = 'CARREGANDO FOTOS...';

        const aboutImageBase64 = await imageToBase64('fotos/image.png');
        const hostingMountainBase64 = await imageToBase64(
            'https://images.unsplash.com/photo-1464822759023-fed622ff2c3b?auto=format&fit=crop&q=80&w=1000'
        );
        const meetingTableBase64 = await imageToBase64(
            'https://i.pinimg.com/1200x/cd/3a/6a/cd3a6a91042667f9f8cd008f8c078222.jpg'
        );

        if (pptxBtnText) pptxBtnText.textContent = 'MONTANDO SLIDES...';

        // ==================================================
        // SLIDE 1: CAPA
        // ==================================================
        let slide = pptx.addSlide();
        slide.background = { color: COLORS.white };

        slide.addShape(pptx.ShapeType.rect, {
            x: 0.5, y: 0.45, w: 0.15, h: 0.15,
            fill: { color: COLORS.black }
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.7, y: 0.45, w: 0.15, h: 0.15,
            fill: { color: COLORS.yellow }
        });
        slide.addText('DESENVOLVIDO POR YAN', {
            x: 1.0, y: 0.35, w: 3.5, h: 0.35,
            fontSize: 10, bold: true, color: COLORS.black,
            fontFace: 'Arial', charSpacing: 2
        });

        // Título na METADE ESQUERDA (0.5 → 5.4)
        slide.addText('PROPOSTA\nCOMERCIAL', {
            x: 0.5, y: 1.7, w: 4.8, h: 2.2,
            fontSize: 44, bold: true, color: COLORS.black,
            fontFace: 'Arial Black', lineSpacingMultiple: 0.95,
            charSpacing: -1.5
        });

        slide.addText('YAN ROCHA\nDEV', {
            x: 0.5, y: 4.5, w: 3, h: 0.6,
            fontSize: 12, bold: true, color: COLORS.black,
            fontFace: 'Arial', lineSpacingMultiple: 1.2,
            charSpacing: 1
        });

        // FOTO à DIREITA (5.5 → 9.7) — usando sizing:cover nativo do PptxGenJS
        const coverBoxX = 5.5;
        const coverBoxY = 0.3;
        const coverBoxW = 4.2;
        const coverBoxH = 5.0;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: coverBoxX, y: coverBoxY, w: coverBoxW, h: coverBoxH,
            fill: { color: COLORS.black },
            rectRadius: 0.3
        });

        if (aboutImageBase64) {
            // ⭐ TRUQUE: usar sizing cover direto — PptxGenJS corta automaticamente
            slide.addImage({
                data: aboutImageBase64,
                x: coverBoxX, y: coverBoxY, w: coverBoxW, h: coverBoxH,
                sizing: { type: 'cover', w: coverBoxW, h: coverBoxH }
            });
        }

        slide.addText('NOME\nEMPRESA', {
            x: coverBoxX, y: coverBoxY + coverBoxH - 1.2, w: coverBoxW, h: 1.2,
            fontSize: 22, bold: true, color: COLORS.white,
            align: 'center', valign: 'middle',
            fontFace: 'Arial Black', lineSpacingMultiple: 1.05
        });

        // ==================================================
        // SLIDE 2: SOBRE MIM
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.yellow };

        // Texto na METADE ESQUERDA (0.5 → 5.7)
        slide.addText('SOBRE MIM', {
            x: 0.5, y: 0.5, w: 4.5, h: 0.8,
            fontSize: 40, bold: true, color: COLORS.black,
            fontFace: 'Arial Black', charSpacing: -2
        });

        slide.addShape(pptx.ShapeType.rect, {
            x: 0.5, y: 1.4, w: 4.5, h: 0.02,
            fill: { color: COLORS.black, transparency: 70 }
        });

        slide.addText('Sou desenvolvedor, formado em Informática pelo ISERJ (FAETEC).', {
            x: 0.5, y: 1.7, w: 4.6, h: 0.5,
            fontSize: 13, bold: true, color: COLORS.black,
            fontFace: 'Arial'
        });

        slide.addText('Fui pesquisador bolsista pela FAPERJ, onde desenvolvi um projeto com HTML, CSS, JS e IA para auxiliar jovens nos vestibulares. Apresentei na Rio Innovation Week, em diversas universidades e na 5ª SNEPT em Brasília. Fui homenageado pela rede FAETEC e premiado pela FAPERJ como um dos melhores pesquisadores Jovens Talentos do Brasil.', {
            x: 0.5, y: 2.3, w: 4.6, h: 3.0,
            fontSize: 10, color: '222222',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        // FOTO à DIREITA (5.8 → 9.7)
        const aboutBoxX = 5.8;
        const aboutBoxY = 0.4;
        const aboutBoxW = 3.9;
        const aboutBoxH = 4.85;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: aboutBoxX, y: aboutBoxY, w: aboutBoxW, h: aboutBoxH,
            fill: { color: COLORS.black },
            rectRadius: 0.3
        });

        if (aboutImageBase64) {
            slide.addImage({
                data: aboutImageBase64,
                x: aboutBoxX, y: aboutBoxY, w: aboutBoxW, h: aboutBoxH,
                sizing: { type: 'cover', w: aboutBoxW, h: aboutBoxH }
            });
        }

        // ==================================================
        // SLIDE 3: ETAPAS DO PROJETO
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.black };

        slide.addText('ETAPAS ', {
            x: 3.0, y: 0.4, w: 3.0, h: 0.7,
            fontSize: 32, bold: true, color: COLORS.white,
            align: 'right', fontFace: 'Arial Black'
        });
        slide.addText('DO PROJETO', {
            x: 6.0, y: 0.4, w: 3.5, h: 0.7,
            fontSize: 32, color: '666666',
            align: 'left', fontFace: 'Arial'
        });

        slide.addShape(pptx.ShapeType.rect, {
            x: 0.5, y: 1.3, w: 9, h: 0.01,
            fill: { color: 'FFFFFF', transparency: 90 }
        });

        const etapas = [
            { num: '01', title: 'ALINHAMENTO\n& ESTRATÉGIA', desc: 'Entendimento completo das necessidades da sua marca e público-alvo.' },
            { num: '02', title: 'PROTÓTIPO\n& DESIGN', desc: 'Criação da interface visual seguindo o alto padrão estético exigido.' },
            { num: '03', title: 'DESENVOLVIMENTO', desc: 'Codificação limpa, responsiva e otimizada para alta performance.' },
            { num: '04', title: 'LANÇAMENTO\n& SUPORTE', desc: 'Publicação oficial, testes rigorosos e acompanhamento contínuo.' }
        ];

        const colWidth = 2.15;
        const colGap = 0.2;
        const startX = 0.5;
        const colY = 1.8;

        etapas.forEach((etapa, i) => {
            const colX = startX + i * (colWidth + colGap);

            slide.addShape(pptx.ShapeType.roundRect, {
                x: colX, y: colY, w: colWidth, h: 3.0,
                fill: { color: '1A1A1A' },
                rectRadius: 0.2,
                line: { color: 'FFFFFF', width: 1, transparency: 90 }
            });

            slide.addText(etapa.num, {
                x: colX + 0.15, y: colY + 0.2, w: colWidth - 0.3, h: 0.8,
                fontSize: 40, bold: true, color: COLORS.yellow,
                fontFace: 'Arial Black'
            });

            slide.addText(etapa.title, {
                x: colX + 0.15, y: colY + 1.05, w: colWidth - 0.3, h: 0.8,
                fontSize: 13, bold: true, color: COLORS.white,
                fontFace: 'Arial Black', lineSpacingMultiple: 1.0
            });

            slide.addText(etapa.desc, {
                x: colX + 0.15, y: colY + 1.9, w: colWidth - 0.3, h: 1.0,
                fontSize: 9, color: '999999',
                fontFace: 'Arial', lineSpacingMultiple: 1.4
            });
        });

        // ==================================================
        // SLIDE 4: PROBLEMA
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.yellow };

        slide.addText('DEV.PROBLEMA', {
            x: 0.5, y: 0.3, w: 3, h: 0.3,
            fontSize: 10, bold: true, color: COLORS.black,
            charSpacing: 1
        });
        slide.addText('02', {
            x: 8.5, y: 0.3, w: 0.5, h: 0.3,
            fontSize: 12, bold: true, color: COLORS.black,
            align: 'right'
        });

        slide.addText('PROBLEMA®', {
            x: 0.3, y: 1.0, w: 9.4, h: 1.5,
            fontSize: 72, bold: true, color: COLORS.white,
            fontFace: 'Arial Black', charSpacing: -4
        });

        slide.addText('Sites lentos, plataformas mal estruturadas e falta de identidade digital que fazem sua empresa perder vendas.', {
            x: 0.5, y: 2.8, w: 9, h: 0.8,
            fontSize: 16, bold: true, color: COLORS.black,
            fontFace: 'Arial', lineSpacingMultiple: 1.3
        });

        slide.addText('A ausência de um design de alto padrão e de código otimizado afasta potenciais clientes e prejudica a credibilidade do seu negócio.', {
            x: 0.5, y: 3.8, w: 9, h: 0.8,
            fontSize: 12, color: '222222',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        // ==================================================
        // SLIDE 5: SOLUÇÃO
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.white };

        slide.addText('DEV.SOLUÇÃO', {
            x: 0.5, y: 0.3, w: 3, h: 0.3,
            fontSize: 10, bold: true, color: COLORS.black,
            charSpacing: 1
        });
        slide.addText('03', {
            x: 8.5, y: 0.3, w: 0.5, h: 0.3,
            fontSize: 12, bold: true, color: COLORS.black,
            align: 'right'
        });

        slide.addText('SOLUÇÃO®', {
            x: 0.3, y: 1.0, w: 9.4, h: 1.5,
            fontSize: 72, bold: true, color: COLORS.yellow,
            fontFace: 'Arial Black', charSpacing: -4
        });

        slide.addText('Desenvolvimento de plataformas digitais de alta performance, combinando design exclusivo.', {
            x: 0.5, y: 2.8, w: 9, h: 0.8,
            fontSize: 16, bold: true, color: COLORS.black,
            fontFace: 'Arial', lineSpacingMultiple: 1.3
        });

        slide.addText('Entregamos soluções rápidas e responsivas focadas na conversão máxima de clientes para o seu negócio.', {
            x: 0.5, y: 3.8, w: 9, h: 0.8,
            fontSize: 12, color: '222222',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        // ==================================================
        // SLIDE 6: FERRAMENTAS
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: 'F7F7F7' };

        slide.addText('Ferramentas\nImportantes', {
            x: 0.5, y: 0.6, w: 5.0, h: 1.6,
            fontSize: 44, bold: false, color: '111111',
            fontFace: 'Arial', lineSpacingMultiple: 0.95
        });

        slide.addText('Para garantir uma operação estratégica e segura, o novo site será entregue com integrações essenciais:', {
            x: 0.5, y: 2.4, w: 4.5, h: 0.6,
            fontSize: 11, color: '555555',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        slide.addText('• Google Analytics: Rastreamento inteligente de métricas para que você possa mensurar o tráfego, entender o comportamento dos visitantes e ter dados reais da sua audiência.\n\n• Conformidade com a LGPD: Implementação de banners de consentimento de cookies e adequação dos formulários, garantindo a proteção de dados e a segurança jurídica.', {
            x: 0.5, y: 3.1, w: 4.5, h: 2.2,
            fontSize: 10, color: '555555',
            fontFace: 'Arial', lineSpacingMultiple: 1.5
        });

        const tools = [
            { num: '[01]', title: 'Google Analytics', desc: 'Rastreamento inteligente de métricas e comportamento de visitantes.', color: COLORS.darkGray, textColor: COLORS.white },
            { num: '[02]', title: 'Conformidade LGPD', desc: 'Banners de consentimento e segurança jurídica total.', color: COLORS.yellow, textColor: '111111' },
            { num: '[03]', title: 'Segurança & Dados', desc: 'Adequação dos formulários e proteção de informações reais.', color: 'E8E8E8', textColor: '111111' }
        ];

        let cardY = 0.8;
        tools.forEach((tool) => {
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 5.5, y: cardY, w: 4.2, h: 1.4,
                fill: { color: tool.color },
                rectRadius: 0.15
            });
            slide.addText(tool.num, {
                x: 5.7, y: cardY + 0.15, w: 0.6, h: 0.3,
                fontSize: 9, bold: true, color: tool.textColor
            });
            slide.addText(tool.title, {
                x: 6.3, y: cardY + 0.15, w: 3.2, h: 0.35,
                fontSize: 15, bold: true, color: tool.textColor,
                fontFace: 'Arial'
            });
            slide.addText(tool.desc, {
                x: 6.3, y: cardY + 0.6, w: 3.2, h: 0.65,
                fontSize: 10, color: tool.textColor,
                fontFace: 'Arial', lineSpacingMultiple: 1.3
            });
            cardY += 1.55;
        });

        // ==================================================
        // SLIDE 7: HOSPEDAGEM — imagem LIMITADA à direita
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.white };

        slide.addText('A', {
            x: 0.5, y: 0.4, w: 0.5, h: 0.7,
            fontSize: 36, bold: true, color: '111111',
            fontFace: 'Arial Black'
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 1.0, y: 0.45, w: 2.8, h: 0.6,
            fill: { color: COLORS.green },
            rectRadius: 0.3
        });
        slide.addText('hospedagem', {
            x: 1.0, y: 0.45, w: 2.8, h: 0.6,
            fontSize: 26, bold: true, color: '111111',
            align: 'center', valign: 'middle',
            fontFace: 'Arial Black'
        });

        const hostingTexts = [
            { bold: 'Hospedagem de Alta Performance:', text: ' Recomendação do plano Premium da Hostinger.' },
            { bold: 'Domínio e E-mail Corporativo:', text: ' Registro do domínio gratuito no 1º ano.' },
            { bold: 'Pronto para o Futuro:', text: ' O plano permite hospedar múltiplos sites.' },
            { bold: 'Titularidade e Transparência:', text: ' A conta e o pagamento ficam vinculados ao CNPJ da Proative.' },
            { bold: 'Suporte Técnico Dedicado:', text: ' Fornecerei auxílio remoto completo.' }
        ];

        let textY = 1.4;
        hostingTexts.forEach((item) => {
            slide.addText([
                { text: item.bold, options: { bold: true, color: '111111' } },
                { text: item.text, options: { color: '444444' } }
            ], {
                x: 0.5, y: textY, w: 4.7, h: 0.6,
                fontSize: 9.5, fontFace: 'Arial', lineSpacingMultiple: 1.35
            });
            textY += 0.72;
        });

        // FOTO da montanha LIMITADA ao canto direito (5.4 → 9.7)
        const hostBoxX = 5.4;
        const hostBoxY = 0.6;
        const hostBoxW = 4.3;
        const hostBoxH = 4.5;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: hostBoxX, y: hostBoxY, w: hostBoxW, h: hostBoxH,
            fill: { color: 'CCCCCC' },
            rectRadius: 0.2
        });

        if (hostingMountainBase64) {
            slide.addImage({
                data: hostingMountainBase64,
                x: hostBoxX, y: hostBoxY, w: hostBoxW, h: hostBoxH,
                sizing: { type: 'cover', w: hostBoxW, h: hostBoxH }
            });
        }

        // Card verde (totalmente dentro do box)
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 7.3, y: 0.5, w: 2.3, h: 1.9,
            fill: { color: COLORS.green },
            rectRadius: 0.2,
            line: { color: COLORS.white, width: 5 }
        });
        slide.addText('🛒', {
            x: 7.5, y: 0.6, w: 0.4, h: 0.3,
            fontSize: 14
        });
        slide.addText('O investimento na plataforma está em promoção por apenas', {
            x: 7.5, y: 0.95, w: 2.0, h: 0.5,
            fontSize: 8, color: '111111',
            fontFace: 'Arial', lineSpacingMultiple: 1.3
        });
        slide.addText('R$ 10,99/mês', {
            x: 7.5, y: 1.5, w: 2.0, h: 0.35,
            fontSize: 15, bold: true, color: '111111',
            fontFace: 'Arial Black'
        });
        slide.addText('no plano de 48 meses.', {
            x: 7.5, y: 1.9, w: 2.0, h: 0.3,
            fontSize: 8, color: '111111',
            fontFace: 'Arial'
        });

        // Card azul (totalmente dentro do box)
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.5, y: 3.3, w: 2.2, h: 1.9,
            fill: { color: COLORS.blue },
            rectRadius: 0.2,
            line: { color: COLORS.white, width: 5 }
        });
        slide.addText('◈', {
            x: 5.7, y: 3.4, w: 0.4, h: 0.3,
            fontSize: 16, color: COLORS.white
        });
        slide.addText('Domínio grátis', {
            x: 5.7, y: 3.9, w: 1.9, h: 0.35,
            fontSize: 13, bold: true, color: COLORS.white,
            fontFace: 'Arial Black'
        });
        slide.addText('no 1º ano e criação de contas de e-mail profissionais.', {
            x: 5.7, y: 4.3, w: 1.9, h: 0.8,
            fontSize: 8, color: 'FFFFFF',
            fontFace: 'Arial', lineSpacingMultiple: 1.3
        });

        // ==================================================
        // SLIDE 8: GARANTIA & MANUTENÇÃO
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.black };

        slide.addText('Garantia &\nManutenção', {
            x: 0.5, y: 0.8, w: 4.5, h: 2,
            fontSize: 44, bold: true, color: COLORS.white,
            fontFace: 'Arial Black', lineSpacingMultiple: 1.0,
            charSpacing: -2
        });

        slide.addText('Nosso compromisso não termina no lançamento. Oferecemos suporte estruturado para que sua plataforma se mantenha estável, segura e sempre atualizada.', {
            x: 0.5, y: 3.0, w: 4.2, h: 1.5,
            fontSize: 13, color: 'A0A0A0',
            fontFace: 'Arial', lineSpacingMultiple: 1.5
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.2, y: 0.8, w: 4.3, h: 2.0,
            fill: { color: COLORS.yellow },
            rectRadius: 0.2
        });
        slide.addText('Garantia de 30 dias', {
            x: 5.5, y: 1.0, w: 3.8, h: 0.5,
            fontSize: 20, bold: true, color: COLORS.black,
            fontFace: 'Arial Black'
        });
        slide.addText('O projeto conta com uma garantia técnica gratuita de 30 dias após a publicação para a correção de qualquer eventual falha, erro de responsividade ou bug estrutural no código.', {
            x: 5.5, y: 1.5, w: 3.8, h: 1.2,
            fontSize: 10, color: '222222',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.2, y: 3.1, w: 4.3, h: 1.8,
            fill: { color: COLORS.yellow },
            rectRadius: 0.2
        });
        slide.addText('Manutenção Avulsa', {
            x: 5.5, y: 3.3, w: 3.8, h: 0.5,
            fontSize: 20, bold: true, color: COLORS.black,
            fontFace: 'Arial Black'
        });
        slide.addText('Alterações após a entrega (textos, imagens, etc.) terão taxa a partir de R$ 100,00 por solicitação.', {
            x: 5.5, y: 3.8, w: 3.8, h: 0.9,
            fontSize: 11, color: '222222',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        // ==================================================
        // SLIDE 9: PRAZO — imagem da reunião LIMITADA ao painel direito
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.navy };

        // Painel superior esquerdo
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.3, y: 0.3, w: 2.9, h: 2.7,
            fill: { color: COLORS.lightGray },
            rectRadius: 0.15
        });
        slide.addText('PRAZO DE ENTREGA', {
            x: 0.5, y: 0.55, w: 2.5, h: 0.25,
            fontSize: 8, bold: true, color: '555555',
            fontFace: 'Arial', charSpacing: 1
        });
        slide.addText('10 Dias', {
            x: 0.5, y: 0.9, w: 2.5, h: 0.9,
            fontSize: 36, bold: true, color: COLORS.navy,
            fontFace: 'Arial Black', charSpacing: -2
        });
        slide.addText('úteis para a entrega do site, após a aprovação do design.', {
            x: 0.5, y: 1.95, w: 2.5, h: 0.8,
            fontSize: 10, color: '444444',
            fontFace: 'Arial', lineSpacingMultiple: 1.3
        });

        // Painel inferior esquerdo
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.3, y: 3.2, w: 2.9, h: 2.2,
            fill: { color: COLORS.lightGray },
            rectRadius: 0.15
        });
        slide.addText('INTRODUÇÃO', {
            x: 0.5, y: 3.45, w: 2.5, h: 0.25,
            fontSize: 8, bold: true, color: '555555',
            fontFace: 'Arial', charSpacing: 1
        });
        slide.addText('Uma visão de performance, agilidade e alta conversão.', {
            x: 0.5, y: 3.75, w: 2.5, h: 0.9,
            fontSize: 12, bold: true, color: COLORS.navy,
            fontFace: 'Arial Black', lineSpacingMultiple: 1.15,
            charSpacing: -0.5
        });
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.5, y: 4.75, w: 1.6, h: 0.4,
            fill: { color: COLORS.navy },
            rectRadius: 0.2
        });
        slide.addText('Começar Projeto', {
            x: 0.5, y: 4.75, w: 1.6, h: 0.4,
            fontSize: 9, bold: true, color: COLORS.white,
            align: 'center', valign: 'middle', fontFace: 'Arial'
        });

        // Painel direito
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 3.4, y: 0.3, w: 6.3, h: 5.1,
            fill: { color: COLORS.lightGray },
            rectRadius: 0.15
        });

        slide.addText('Construímos o seu\nfuturo digital.', {
            x: 3.7, y: 0.55, w: 5.8, h: 1.3,
            fontSize: 32, bold: true, color: COLORS.navy,
            fontFace: 'Arial Black', lineSpacingMultiple: 0.95,
            charSpacing: -2
        });

        // IMAGEM da reunião — limitada ao box (3.7 → 9.4) x (2.0 → 5.1)
        const meetBoxX = 3.7;
        const meetBoxY = 2.0;
        const meetBoxW = 5.7;
        const meetBoxH = 3.1;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: meetBoxX, y: meetBoxY, w: meetBoxW, h: meetBoxH,
            fill: { color: '222222' },
            rectRadius: 0.15
        });

        if (meetingTableBase64) {
            slide.addImage({
                data: meetingTableBase64,
                x: meetBoxX, y: meetBoxY, w: meetBoxW, h: meetBoxH,
                sizing: { type: 'cover', w: meetBoxW, h: meetBoxH }
            });
        } else {
            slide.addText('[Imagem: Reunião]', {
                x: meetBoxX, y: meetBoxY, w: meetBoxW, h: meetBoxH,
                fontSize: 12, color: '666666',
                align: 'center', valign: 'middle',
                fontFace: 'Arial'
            });
        }

        // ==================================================
        // SLIDE 10: VALORES
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.black };

        slide.addText([
            { text: 'Investimento transparente', options: { color: '888888', bold: true } },
            { text: ' — e facilitado', options: { color: COLORS.white, bold: true } }
        ], {
            x: 0.5, y: 0.4, w: 9, h: 0.6,
            fontSize: 24, fontFace: 'Arial'
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.3, y: 1.2, w: 4.4, h: 3.8,
            fill: { color: 'F4F5F5' },
            rectRadius: 0.2
        });
        slide.addText('Hostinger (Hospedagem)', {
            x: 0.6, y: 1.5, w: 3.8, h: 0.5,
            fontSize: 18, bold: true, color: '111111',
            fontFace: 'Arial Black'
        });
        slide.addText('Plano: Premium (com domínio grátis no 1º ano).\n\nValor Promocional: R$ 10,99/mês (no ciclo de 48 meses).\n\nA contratação e o pagamento desta etapa são feitos diretamente pela Proative na plataforma da Hostinger.', {
            x: 0.6, y: 2.1, w: 3.8, h: 2.0,
            fontSize: 10, color: '555555',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.6, y: 4.2, w: 1.3, h: 0.35,
            fill: { color: COLORS.white },
            rectRadius: 0.17,
            line: { color: 'E0E0E0', width: 1 }
        });
        slide.addText('Domínio Grátis', {
            x: 0.6, y: 4.2, w: 1.3, h: 0.35,
            fontSize: 8, bold: true, color: '333333',
            align: 'center', valign: 'middle'
        });
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 2.0, y: 4.2, w: 1.3, h: 0.35,
            fill: { color: 'E6F6EB' },
            rectRadius: 0.17,
            line: { color: 'D1ECD9', width: 1 }
        });
        slide.addText('Plano Premium', {
            x: 2.0, y: 4.2, w: 1.3, h: 0.35,
            fontSize: 8, bold: true, color: '166534',
            align: 'center', valign: 'middle'
        });
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 3.4, y: 4.2, w: 1.3, h: 0.35,
            fill: { color: COLORS.white },
            rectRadius: 0.17,
            line: { color: 'E0E0E0', width: 1 }
        });
        slide.addText('Certificado SSL', {
            x: 3.4, y: 4.2, w: 1.3, h: 0.35,
            fontSize: 8, bold: true, color: '333333',
            align: 'center', valign: 'middle'
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.0, y: 1.2, w: 4.5, h: 3.8,
            fill: { color: '222222' },
            rectRadius: 0.2
        });
        slide.addText('Desenvolvimento', {
            x: 5.3, y: 1.5, w: 4.0, h: 0.5,
            fontSize: 18, bold: true, color: COLORS.white,
            fontFace: 'Arial Black'
        });

        slide.addText('Investimento Total', {
            x: 5.3, y: 2.2, w: 3.0, h: 0.3,
            fontSize: 12, color: 'D1D5DB', fontFace: 'Arial'
        });
        slide.addText('R$ 1.000,00', {
            x: 8.0, y: 2.2, w: 1.3, h: 0.3,
            fontSize: 14, bold: true, color: COLORS.white,
            align: 'right', fontFace: 'Arial'
        });

        slide.addShape(pptx.ShapeType.rect, {
            x: 5.3, y: 2.7, w: 4.0, h: 0.01,
            fill: { color: 'FFFFFF', transparency: 90 }
        });

        slide.addText('Sinal (Entrada)', {
            x: 5.3, y: 2.9, w: 2.5, h: 0.3,
            fontSize: 12, color: 'D1D5DB', fontFace: 'Arial'
        });
        slide.addText('R$ 200,00', {
            x: 8.0, y: 2.9, w: 1.3, h: 0.3,
            fontSize: 14, bold: true, color: COLORS.white,
            align: 'right', fontFace: 'Arial'
        });
        slide.addText('Reserva de agenda e início imediato.', {
            x: 5.3, y: 3.2, w: 2.5, h: 0.3,
            fontSize: 8, color: '888888', fontFace: 'Arial'
        });

        slide.addText('Restante (Entrega)', {
            x: 5.3, y: 3.6, w: 2.5, h: 0.3,
            fontSize: 12, color: 'D1D5DB', fontFace: 'Arial'
        });
        slide.addText('R$ 800,00', {
            x: 8.0, y: 3.6, w: 1.3, h: 0.3,
            fontSize: 14, bold: true, color: COLORS.white,
            align: 'right', fontFace: 'Arial'
        });
        slide.addText('Pagos mediante publicação do projeto.', {
            x: 5.3, y: 3.9, w: 2.5, h: 0.3,
            fontSize: 8, color: '888888', fontFace: 'Arial'
        });

        slide.addText('Parcelamento flexível: Dividimos o valor do projeto no cartão de crédito, ou via Pix e Boleto.', {
            x: 5.3, y: 4.4, w: 4.0, h: 0.4,
            fontSize: 9, color: '9CA3AF',
            fontFace: 'Arial', lineSpacingMultiple: 1.3
        });

        // ==================================================
        // SLIDE 11: CONTATO
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: '121212' };

        slide.addText('FALE COMIGO', {
            x: 0.5, y: 0.5, w: 9, h: 0.8,
            fontSize: 44, bold: true, color: COLORS.white,
            fontFace: 'Arial Black', charSpacing: -2
        });
        slide.addText('Tem algum projeto em mente, dúvida sobre valores ou prazos? Vamos conversar e tirar sua ideia do papel!', {
            x: 0.5, y: 1.3, w: 9, h: 0.5,
            fontSize: 13, color: '9CA3AF',
            fontFace: 'Arial'
        });

        slide.addText('E-MAIL DIRETO', {
            x: 0.5, y: 2.2, w: 4, h: 0.3,
            fontSize: 12, bold: true, color: COLORS.yellow,
            charSpacing: 2
        });
        slide.addText('Prefere enviar uma mensagem detalhada com os requisitos do seu projeto ou dúvidas técnicas? Mande um e-mail e responderei rapidamente.', {
            x: 0.5, y: 2.6, w: 4.2, h: 0.8,
            fontSize: 11, color: '9CA3AF',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });
        slide.addText('YANROCHADEV@GMAIL.COM', {
            x: 0.5, y: 3.6, w: 4.2, h: 0.5,
            fontSize: 18, bold: true, color: COLORS.white,
            fontFace: 'Arial Black'
        });

        slide.addText('INICIAR PROJETO', {
            x: 5.5, y: 2.2, w: 4, h: 0.3,
            fontSize: 12, bold: true, color: COLORS.yellow,
            charSpacing: 2
        });
        slide.addText('Pronto para transformar sua presença digital com uma plataforma de alto impacto? Entre em contato agora e garanta sua vaga na agenda.', {
            x: 5.5, y: 2.6, w: 4.2, h: 0.8,
            fontSize: 11, color: '9CA3AF',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.5, y: 3.6, w: 2.2, h: 0.45,
            fill: { color: COLORS.yellow },
            rectRadius: 0.22
        });
        slide.addText('Enviar e-mail →', {
            x: 5.5, y: 3.6, w: 2.2, h: 0.45,
            fontSize: 11, bold: true, color: COLORS.black,
            align: 'center', valign: 'middle', fontFace: 'Arial'
        });

        slide.addText('WHATSAPP / SUPORTE', {
            x: 5.5, y: 4.3, w: 4, h: 0.3,
            fontSize: 12, bold: true, color: COLORS.yellow,
            charSpacing: 2
        });
        slide.addText('Tem alguma dúvida urgente? Fale diretamente comigo pelo WhatsApp para um atendimento imediato.', {
            x: 5.5, y: 4.6, w: 4.2, h: 0.5,
            fontSize: 10, color: '9CA3AF',
            fontFace: 'Arial', lineSpacingMultiple: 1.4
        });

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.5, y: 5.1, w: 2.4, h: 0.45,
            fill: { color: '1F1F1F' },
            rectRadius: 0.22,
            line: { color: 'FFFFFF', width: 1, transparency: 90 }
        });
        slide.addText('Chamar no WhatsApp', {
            x: 5.5, y: 5.1, w: 2.4, h: 0.45,
            fontSize: 10, bold: true, color: COLORS.white,
            align: 'center', valign: 'middle', fontFace: 'Arial'
        });

        // ==================================================
        // SLIDE 12: OBRIGADO
        // ==================================================
        slide = pptx.addSlide();
        slide.background = { color: COLORS.black };

        slide.addText('YanDEV', {
            x: 0, y: 1.5, w: 10, h: 1.5,
            fontSize: 72, bold: true, color: COLORS.yellow,
            align: 'center', valign: 'middle',
            fontFace: 'Arial Black'
        });
        slide.addText('Obrigado pela atenção!', {
            x: 0, y: 3.2, w: 10, h: 0.8,
            fontSize: 24, color: '888888',
            align: 'center', valign: 'middle',
            fontFace: 'Arial'
        });
        slide.addText('yanrochadev@gmail.com', {
            x: 0, y: 4.2, w: 10, h: 0.5,
            fontSize: 14, color: '666666',
            align: 'center', valign: 'middle',
            fontFace: 'Arial'
        });

        // SALVAR
        await pptx.writeFile({ fileName: 'proposta-comercial-yandev.pptx' });
    }

    // --- FUNCIONALIDADES DA PÁGINA ---
    const toggleBtn = document.getElementById('toggleThemeBtn');
    const wakeupSection = document.getElementById('desenvolvimento');
    
    if (toggleBtn && wakeupSection) {
        const problemState = wakeupSection.querySelector('.problem-state');
        const solutionState = wakeupSection.querySelector('.solution-state');
        const svgLetterText = document.getElementById('svgLetterText');
        const sectionModeTitle = document.getElementById('sectionModeTitle');
        const toggleBtnText = document.getElementById('toggleBtnText');

        toggleBtn.addEventListener('click', () => {
            const isSolution = wakeupSection.classList.toggle('is-solution');
            
            if (isSolution) {
                problemState.classList.remove('active');
                solutionState.classList.add('active');
                svgLetterText.textContent = 'SOLUÇÃO®';
                sectionModeTitle.textContent = 'DEV.SOLUÇÃO';
                toggleBtnText.textContent = 'Ver problema';
            } else {
                solutionState.classList.remove('active');
                problemState.classList.add('active');
                svgLetterText.textContent = 'PROBLEMA®';
                sectionModeTitle.textContent = 'DEV.PROBLEMA';
                toggleBtnText.textContent = 'Ver a solução';
            }
        });
    }

    const payTabs = document.querySelectorAll('.sys-tab');
    payTabs.forEach(tab => {
        tab.addEventListener('click', function() {
            payTabs.forEach(t => {
                t.classList.remove('active');
                const dot = t.querySelector('.sys-dot');
                if(dot) dot.remove();
            });

            this.classList.add('active');
            const newDot = document.createElement('span');
            newDot.classList.add('sys-dot');
            this.prepend(newDot);
        });
    });

    const faqItems = document.querySelectorAll('.faq-item');
    faqItems.forEach(item => {
        const header = item.querySelector('.faq-header');
        header.addEventListener('click', () => {
            const isActive = item.classList.contains('active');
            faqItems.forEach(i => {
                i.classList.remove('active');
                const body = i.querySelector('.faq-body');
                if(body) body.style.maxHeight = null;
            });

            if (!isActive) {
                item.classList.add('active');
                const body = item.querySelector('.faq-body');
                if(body) body.style.maxHeight = body.scrollHeight + 30 + 'px';
            }
        });
    });
});