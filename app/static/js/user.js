const typeInputs = document.querySelectorAll('input[name="tipo_entrega"]');
const personalFields = document.querySelectorAll('.personal-only');
const envioFields = document.querySelectorAll('.envio-only');
const moreProducts = document.querySelectorAll('input[name="mas_productos"]');
const extraWrap = document.getElementById('extra-codes-wrap');
const addCodeBtn = document.getElementById('add-code-btn');
const extraCodes = document.getElementById('extra-codes');
const codeInputs = [
    document.querySelector('input[name="codigo"]'),
    document.querySelector('input[name="codigo_principal"]'),
].filter(Boolean);

function formatCode(rawValue) {
    const clean = (rawValue || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 5);
    if (!clean) return '';

    if (clean.length <= 2) {
        return clean.length === 2 ? `${clean}-` : clean;
    }

    if (clean.length <= 4) {
        const left = clean.slice(0, 2);
        const middle = clean.slice(2);
        return clean.length === 4 ? `${left}-${middle}-` : `${left}-${middle}`;
    }

    const left = clean.slice(0, 2);
    const middle = clean.slice(2, 4);
    const right = clean.slice(4, 5);
    return `${left}-${middle}-${right}`;
}

function isCompleteCode(value) {
    return /^[A-Z]{2}-[0-9]{2}-[A-Z]$/.test(value || '');
}

function syncCodeFeedback(input) {
    if (!input) return;

    const helper = input.closest('form')?.querySelector('[data-code-feedback]');
    const value = input.value.trim();
    const complete = isCompleteCode(value);

    input.classList.toggle('invalid-code', Boolean(value) && !complete);

    if (helper) {
        helper.textContent = !value
            ? 'Formato esperado: AA-00-A'
            : complete
                ? 'Codigo valido'
                : 'Se completa solo: mayusculas y guiones automaticos';
        helper.classList.toggle('is-valid', complete);
    }
}

function bindFormattedCodeInput(input) {
    if (!input) return;

    const applyFormat = () => {
        const nextValue = formatCode(input.value);
        if (input.value !== nextValue) {
            input.value = nextValue;
        }
        syncCodeFeedback(input);
    };

    input.addEventListener('input', applyFormat);
    input.addEventListener('blur', applyFormat);
    applyFormat();
}

function toggleDeliveryFields() {
    const selected = document.querySelector('input[name="tipo_entrega"]:checked');
    const type = selected ? selected.value : '';

    personalFields.forEach((field) => {
        field.classList.toggle('hidden', type !== 'Entrega personal');
        const input = field.querySelector('input, select');
        if (input) input.required = type === 'Entrega personal';
    });

    envioFields.forEach((field) => {
        field.classList.toggle('hidden', type !== 'Envio');
        const input = field.querySelector('input, select');
        if (input && input.name !== 'direccion_referencia') input.required = type === 'Envio';
    });
}

function addCodeInput() {
    const wrapper = document.createElement('div');
    wrapper.className = 'inline-form';

    const input = document.createElement('input');
    input.name = 'codigos_extra';
    input.placeholder = 'Codigo adicional';
    input.maxLength = 7;

    input.addEventListener('input', () => {
        const nextValue = formatCode(input.value);
        if (input.value !== nextValue) {
            input.value = nextValue;
        }
    });

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-light';
    removeBtn.textContent = 'Quitar';
    removeBtn.addEventListener('click', () => wrapper.remove());

    wrapper.appendChild(input);
    wrapper.appendChild(removeBtn);
    extraCodes.appendChild(wrapper);
}

function toggleExtraCodes() {
    const selected = document.querySelector('input[name="mas_productos"]:checked');
    const show = selected && selected.value === 'si';
    extraWrap.classList.toggle('hidden', !show);

    if (show && extraCodes.children.length === 0) {
        addCodeInput();
    }

    if (!show) {
        extraCodes.innerHTML = '';
    }
}

if (typeInputs.length) {
    typeInputs.forEach((input) => input.addEventListener('change', toggleDeliveryFields));
    toggleDeliveryFields();
}

if (moreProducts.length) {
    moreProducts.forEach((input) => input.addEventListener('change', toggleExtraCodes));
    toggleExtraCodes();
}

if (addCodeBtn) {
    addCodeBtn.addEventListener('click', addCodeInput);
}

codeInputs.forEach(bindFormattedCodeInput);

extraCodes?.addEventListener('input', (event) => {
    const target = event.target;
    if (target instanceof HTMLInputElement && target.name === 'codigos_extra') {
        const nextValue = formatCode(target.value);
        if (target.value !== nextValue) {
            target.value = nextValue;
        }
    }
});
