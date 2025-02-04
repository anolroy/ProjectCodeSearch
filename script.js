document.addEventListener('DOMContentLoaded', () => {
    const searchInput = document.getElementById('search');
    const resultsContainer = document.getElementById('results');
    let codeData = {};

    // Load JSON data
    fetch('\output.json')
       // .then(response => response.json())
        .then(response => response.json())
         
        .then(data => codeData = data)
        .catch(error => console.error('Error loading JSON data:', error));

        searchInput.addEventListener('keyup', () => {
            const query = searchInput.value.toLowerCase();
            resultsContainer.innerHTML = '';
        
            if (query) {
                // Filter codeData keys based on the query
                const results = Object.keys(codeData)
                    .filter(key => key.toLowerCase().includes(query))
                    .map(key => ({ name: key, code: codeData[key] }));
        
                if (results.length > 0) {
                    // Display each matching snippet
                    results.forEach(result => {
                        const snippetDiv = document.createElement('div');
                        snippetDiv.classList.add('snippet');
                        snippetDiv.innerHTML = `
                            <h3>${result.name}</h3>
                            <pre><code>${result.code}</code></pre>
                        `;
                        resultsContainer.appendChild(snippetDiv);
                    });
                } else {
                    // Display message if no matches are found
                    resultsContainer.innerHTML = '<p>No matching code snippets found.</p>';
                }
            }
        });
});
// Updated: Receipts:
// 	receipt with invoice type=1
// 	due receipt type=5

// Refunds:
// orderheader table will be used for refund line item and type shall be=2 
// Updated transaction type:

// 			if ($hasRefundLineItems ) 
//                         {
//                             $TransactionType = 4;
//                             $TransactionTypeDesc="Refund with item lines";
//                         }
//                     elseif ($hasRefundLineItems)
//                         {
//                             $TransactionType = 3;
//                             $TransactionTypeDesc="Sales return Refund";
//                         }