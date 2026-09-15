<?php

namespace App\Services;

use Illuminate\Support\Collection;
use Illuminate\Session\SessionManager;
use Illuminate\Support\Facades\Log;
use App\Models\TempCart;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Auth;
use App\Repositories\ProfitRepository;
use Exception;
use Carbon\Carbon;
use App\Models\Pedido;
use App\Models\PedidoItem;
use Illuminate\Support\Str;
use Illuminate\Support\Facades\Http; // <-- AQUI ESTA LA LINEA FALTANTE
use Illuminate\Support\Facades\Cache;


class CartService
{
    protected $session;
    protected $profitRepository;

    public function __construct(SessionManager $session, ProfitRepository $profitRepository)
    {
        $this->session = $session;
        $this->profitRepository = $profitRepository;

    }

    /**
     * Helper para obtener el identificador único del carrito (usuario o sesión).
     *
     * @return array
     */

    private function getCartIdentifier(): array
    {
        // 1. Prioridad: Revisar la sesión personalizada (Donde vive FAR00034)
        // Según tu log de DEBUG, el sistema obtiene el ID de la sesión
        if (session()->has('codigo_cliente')) { // <--- ASEGÚRATE QUE EL NOMBRE SEA 'codigo_cliente' o como lo tengas en tu login
            return [
                'column' => 'user_id',
                'value'  => trim(session()->get('codigo_cliente')),
            ];
        }

        // 2. Fallback: Auth estándar
        if (auth()->check()) {
            return [
                'column' => 'user_id',
                'value'  => auth()->id(),
            ];
        }

        // 3. Invitado: UUID
        if (!session()->has('cart_session_id')) {
            session()->put('cart_session_id', (string) \Illuminate\Support\Str::uuid());
        }

        return [
            'column' => 'session_id',
            'value'  => session()->get('cart_session_id'),
        ];
    }

    /**
     * Obtiene los ítems del carrito del usuario actual desde la base de datos.
     *
     * @return \Illuminate\Support\Collection
     */

    public function content(): Collection
    {
        // --- 1. Obtención de Items del Carrito ---
        $identifier = TempCart::getCartIdentifier();

        $cartItems = TempCart::where($identifier['column'], $identifier['value'])
            ->where('status', 0)
            ->select('id', 'product_id', 'product_name', 'price', 'quantity', 'options')
            ->get();

        if ($cartItems->isEmpty()) {
            return new Collection();
        }

        // --- 2. Cliente y almacenes (cacheado) ---
        $coCli = Auth::check() ? Auth::user()->co_cli : null;
        $primaryWarehouse   = '01';
        $secondaryWarehouse = '04';

        if ($coCli) {
            $client = Cache::remember("cart_client_data:{$coCli}", now()->addMinutes(10), function () use ($coCli) {
                return $this->profitRepository->getClienteByCoCli($coCli);
            });

            if (trim((string)($client['fax'] ?? '')) === '02') {
                [$primaryWarehouse, $secondaryWarehouse] = [$secondaryWarehouse, $primaryWarehouse];
            }
        }

        // --- 3. Productos cacheados en una sola llamada ---
        $productIds = $cartItems->pluck('product_id')->unique()->toArray();
        
        // Clave de caché basada en los IDs + cliente para evitar mezclar precios
        $cacheKey = 'cart_products:' . md5(implode(',', $productIds) . ($coCli ?? 'guest'));
        
        $profitMap = Cache::remember($cacheKey, now()->addMinutes(5), function () use ($productIds, $coCli) {
            return $this->getCachedProfitProducts($productIds, $coCli);
        });

        // --- 4. Parsear options UNA sola vez fuera del map ---
        $optionsParsed = $cartItems->keyBy('id')->map(function ($item) {
            return is_array($item->options)
                ? $item->options
                : (json_decode($item->options ?? '[]', true) ?? []);
        });

        // --- 5. Mapeo final (sin queries adentro) ---
        return $cartItems->map(function ($item) use ($profitMap, $primaryWarehouse, $secondaryWarehouse, $optionsParsed) {
            $profit  = $profitMap[trim($item->product_id)] ?? null;
            $options = $optionsParsed[$item->id] ?? [];

            $price   = (float) $item->price;
            $name    = trim($item->product_name);
            $code    = trim($item->product_id);
            $tipoImp = 'N';
            $stockSC = 0.0;
            $stockBQ = 0.0;

            if ($profit) {
                // No reemplazar el precio almacenado: ya tiene los descuentos por art/lin/cat/prov aplicados.
                // Sobreescribirlo con el precio crudo de Profit eliminaría esos descuentos.
                $name    = trim($profit['des_art'] ?? $name);
                $code    = trim($profit['co_art']  ?? $code);
                $tipoImp = trim($profit['tipo_imp'] ?? 'N');
                $stockSC = (float)($profit['stock_sc'] ?? 0.0);
                $stockBQ = (float)($profit['stock_bq'] ?? 0.0);
            }

            $stockPrimary   = $primaryWarehouse === '01' ? $stockSC : $stockBQ;
            $stockSecondary = $primaryWarehouse === '01' ? $stockBQ : $stockSC;
            $cantidadTotal  = (int) $item->quantity;

            return [
                'rowId'                        => (int) $item->id,
                'co_art'                       => $code,
                'des_art'                      => $name,
                'total_art'                    => $cantidadTotal,
                'prec_agr'                     => round($price, 2),
                'subtotal'                     => round($price * $cantidadTotal, 2),
                'options'                      => $options,
                'stock_sc'                     => $stockSC,
                'stock_bq'                     => $stockBQ,
                'primary_warehouse_name'       => $primaryWarehouse,
                'secondary_warehouse_name'     => $secondaryWarehouse,
                'stock_in_primary_warehouse'   => $stockPrimary,
                'stock_in_secondary_warehouse' => $stockSecondary,
                'tipo_imp'                     => $tipoImp,
            ];
        })
        ->filter(fn($item) => !empty($item['co_art']))
        ->values();
    }

    /**
     * ⚡ FUNCIÓN AUXILIAR CRÍTICA PARA EL CACHEO DE PRODUCTOS ⚡
     * Implementar esto en CartService para evitar llamadas Profit lentas.
     */

    protected function getCachedProfitProducts(array $productIds, ?string $coCli): Collection
    {
        // Construye una clave de caché única basada en los IDs y el cliente
        $sortedIds = implode(',', $productIds);
        $cacheKey = "profit_products:cli_{$coCli}:ids_" . sha1($sortedIds);
        
        // TTL (Time To Live) de 60 segundos: el stock y precio se actualizarán al menos cada minuto.
        $ttlSeconds = 60; 

        return Cache::remember($cacheKey, $ttlSeconds, function () use ($productIds, $coCli) {
            // Esta línea es la única que llama al repositorio lento
            $profitProducts = $this->profitRepository->getProductsByCoArt($productIds, $coCli);
            return $profitProducts->keyBy('co_art');
        });
    }

    /**
     * Agrega un producto al carrito.
     *
     * @param string $productId El ID del producto.
     * @param int $quantity La cantidad del producto a añadir.
     * @param float $price El precio unitario del producto en el momento de la adición.
     * @param array $options Opciones adicionales del producto (opcional, puede contener 'name' e 'image').
     * @return \App\Models\TempCart|false El item del carrito creado o false si falla.
     */

    public function add(string $productId, int $quantity, float $price, array $options = [])
    {
        try {
            $identifier = TempCart::getCartIdentifier();

            $name = trim($options['name'] ?? $productId);

            $existingItem = TempCart::where($identifier['column'], $identifier['value'])
                ->where('product_id', $productId)
                ->where('status', 0)
                ->first();

            if ($existingItem) {
                $existingItem->quantity += $quantity;
                $existingItem->price     = $price;
                $existingItem->save();
                return $existingItem;
            }

            // Usar $identifier['column'] para soportar tanto user_id como session_id
            return TempCart::create([
                $identifier['column'] => $identifier['value'],
                'product_id'          => $productId,
                'quantity'            => $quantity,
                'price'               => $price,
                'product_name'        => $name,
                'options'             => $options, // array — el cast de TempCart serializa automáticamente
                'status'              => 0,
            ]);
        } catch (\Exception $e) {
            Log::error("CartService::add producto {$productId}: " . $e->getMessage());
            return false;
        }
    }

    /**
     * Calcula el subtotal total del carrito (suma de subtotales de productos).
     *
     * @return float
     */

    public function subtotal(): float
    {
        // Obtener los ítems del carrito (ya procesados con precios, etc.)
        $cartItems = $this->content();

        // Sumar el 'subtotal' de cada ítem
        $subtotal = $cartItems->sum('subtotal');

        return round($subtotal, 2);
    }

    /**
     * Calcula el total final del carrito después de aplicar descuentos (si aplica).
     *
     * @return float
     */

    public function total(): float
    {
        $subtotal = $this->subtotal();

        try {
            $descuentoGlobalCliente = 0;
            if (Auth::check()) {
                $coCli  = Auth::user()->co_cli;
                $client = Cache::remember("cart_client_data:{$coCli}", now()->addMinutes(10), function () use ($coCli) {
                    return $this->profitRepository->getClienteByCoCli($coCli);
                });
                $descuentoGlobalCliente = (float) ($client['desc_glob'] ?? 0);
            }

            if ($descuentoGlobalCliente > 0) {
                return round($subtotal * (1 - $descuentoGlobalCliente / 100), 2);
            }
        } catch (\Exception $e) {
            Log::warning("CartService::total no pudo obtener descuento: " . $e->getMessage());
        }

        return round($subtotal, 2);
    }

    /**
     * Calcula el total del IVA del carrito.
     * Asume un IVA del 16% (ajusta según tu lógica).
     *
     * @return float
     */

    public function totalIva(): float
    {
        $totalIvaAmount = 0;
        $identifier = TempCart::getCartIdentifier();

        $cartItems = TempCart::where($identifier['column'], $identifier['value'])
                             ->where('status', 0)
                             ->get();

        $productIds = $cartItems->pluck('product_id')->unique()->toArray();

        $profitProductsBasicInfo = collect();
        if (!empty($productIds)) {
            $profitProductsBasicInfo = $this->profitRepository->getProductsByCoArt($productIds);
        }

        foreach ($cartItems as $item) {
            $profitProduct = $profitProductsBasicInfo->firstWhere('co_art', $item->product_id);

            // Se ha cambiado a array si getProductsByCoArt devuelve un array asociativo
            if (!$profitProduct || !isset($profitProduct['precio'])) {
                continue; // Saltar este ítem
            }

            $basePrice = (float) $profitProduct['precio']; // Acceder como array
            $quantity = (float) $item->quantity; // Usar la cantidad del ítem del carrito

            // ---- Lógica para inferir IVA de tipo_imp ----
            $ivaRate = 0;
            if (isset($profitProduct['tipo_imp']) && trim($profitProduct['tipo_imp']) === '1') { // Acceder como array
                $ivaRate = 16; // Tasa de IVA para artículos gravados (tipo_imp = '1'). CONFIRMA ESTE VALOR.
            }
            // -------------------------------------------------------------

            $totalItemBasePrice = $basePrice * $quantity;
            $ivaAmountForItem = $totalItemBasePrice * ($ivaRate / 100);

            $totalIvaAmount += $ivaAmountForItem;
        }

        return round($totalIvaAmount, 2);
    }

    /**
     * Elimina todos los items del carrito con status 0 para el identificador actual.
     *
     * @return void
     */

     public function clear(): bool
    {
        try {
            $identifier = TempCart::getCartIdentifier();

            TempCart::where('status', 0)
                ->where($identifier['column'], $identifier['value'])
                ->update(['status' => 2]);

            return true;
        } catch (\Exception $e) {
            Log::error("Error al vaciar: " . $e->getMessage());
            return false;
        }
    }

    // public function clear(): bool
    // {
    //     try {
    //         $identifier = TempCart::getCartIdentifier();

    //         TempCart::where('status', 0)
    //             ->where($identifier['column'], $identifier['value'])
    //             ->delete();

    //         return true;
    //     } catch (\Exception $e) {
    //         Log::error("Error al vaciar: " . $e->getMessage());
    //         return false;
    //     }
    // }

    /**
     * Elimina un producto del carrito basado en su rowId (ID de la tabla temp_cart).
     *
     * @param string $rowId El ID del registro en la tabla 'temp_cart' a eliminar.
     * @return bool True si se eliminó con éxito, false en caso contrario.
     */

    public function remove(string $rowId): bool
    {
        try {
            // Obtener el identificador del carrito actual (usuario_id o session_id)
            $identifier = TempCart::getCartIdentifier();

            // Buscar y eliminar el ítem de la tabla temp_cart
            $deletedCount = TempCart::where($identifier['column'], $identifier['value'])
                                     ->where('id', $rowId) // Usa 'id' que es tu rowId
                                     ->where('status', 0) // Solo eliminar ítems activos en el carrito
                                     ->delete();

            if ($deletedCount > 0) {
                return true;
            } else {
                return false;
            }
        } catch (\Exception $e) {
            Log::error("Error al eliminar item del carrito con rowId {$rowId}: " . $e->getMessage());
            return false;
        }
    }

    /**
     * Calcula el número total de items únicos en el carrito.
     * (Este método es llamado en CartController y debería existir)
     * @return int
     */

    public function count(): int
    {
        $identifier = TempCart::getCartIdentifier();
        return TempCart::where($identifier['column'], $identifier['value'])
                         ->where('status', 0)
                         ->count();
    }

    /**
     * Actualiza la cantidad de un producto en el carrito por su rowId.
     * (Este método es llamado desde CartController@update, debe existir)
     * @param string $rowId El ID del registro en la tabla 'temp_cart'.
     * @param int $newQuantity La nueva cantidad para el producto.
     * @return bool True si se actualizó con éxito, false en caso contrario.
     */

    public function update(string $rowId, int $newQuantity, array $options): bool
    {
        try {
            $identifier = TempCart::getCartIdentifier();

            $item = TempCart::where($identifier['column'], $identifier['value'])
                                ->where('id', $rowId)
                                ->where('status', 0)
                                ->first();

            if ($item) {
                // ✅ $item->options es un ARRAY gracias al cast. Si es NULL, usa array vacío.
                $currentOptions = $item->options ?? []; 
                
                // Si el cast fallara o la columna fuera null, lo aseguramos.
                if (!is_array($currentOptions)) {
                    $currentOptions = [];
                }
                
                // Actualiza las cantidades dentro del array de opciones
                $currentOptions['quantitySC'] = $options['quantitySC'] ?? 0;
                $currentOptions['quantityBQ'] = $options['quantityBQ'] ?? 0;

                // ⛔️ ELIMINAR json_encode: Laravel se encarga de codificar el array a JSON
                //    al guardar, gracias al protected $casts.
                $item->options = $currentOptions; 

                // Actualiza la cantidad total
                $item->quantity = $newQuantity;

                // Guarda los cambios en el modelo
                $item->save();
                return true;
            }

            return false;
        } catch (\Exception $e) {
            // Asegúrate de que este log esté capturando el error:
            Log::error("Error al actualizar item con rowId {$rowId}: " . $e->getMessage(), ['trace' => $e->getTraceAsString()]);
            return false;
        }
    }

    /**
     * Procesa el carrito de compras, lo transforma en la estructura
     * necesaria para un endpoint, guarda en la base de datos local y luego lo envía a una API.
     *
     * @param bool $clearCartAfterProcess Si es verdadero, vacía el carrito después del procesamiento.
     * @return bool Retorna verdadero si el procesamiento fue exitoso.
     * @throws Exception Si ocurre algún error durante el proceso.
     */

    public function processCart(bool $clearCartAfterProcess = true, ?Collection $prefetchedItems = null): bool
    {
        $user = Auth::user();
        if (!$user) {
            Log::error("ProcessCart: usuario no autenticado.");
            throw new Exception("No se puede procesar el carrito: usuario no autenticado.");
        }

        // Reutilizar items ya obtenidos en sendCart para evitar llamar content() otra vez
        $cartItems  = $prefetchedItems ?? $this->content();
        if ($cartItems->isEmpty()) return false;

        $validItems = $cartItems->filter(fn($item) => isset($item['co_art']));
        if ($validItems->isEmpty()) return false;

        // ── 1. Guardar pedido en DB (operación local, rápida) ─────────────────
        $pedido = DB::transaction(function () use ($user, $validItems) {
            Pedido::where('cod_cliente', $user->co_cli)->where('status', 2)->update(['status' => 3]);

            $pedido = Pedido::create(['cod_cliente' => $user->co_cli, 'status' => 2]);
            if (!$pedido || !$pedido->id) throw new Exception("Fallo al crear el pedido en base de datos.");

            $now = Carbon::now();
            PedidoItem::insert($validItems->map(fn($item) => [
                'pedido_id'                => $pedido->id,
                'co_art'                   => $item['co_art'],
                'cant_producto'            => $item['total_art'] ?? 0,
                'precio'                   => $item['prec_agr'] ?? 0,
                'quantitySC'               => (int) ($item['options']['quantitySC'] ?? 0),
                'quantityBQ'               => (int) ($item['options']['quantityBQ'] ?? 0),
                'primary_warehouse_name'   => $item['primary_warehouse_name'] ?? null,
                'secondary_warehouse_name' => $item['secondary_warehouse_name'] ?? null,
                'created_at'               => $now,
                'updated_at'               => $now,
            ])->values()->toArray());

            return $pedido;
        });

        // ── 2. Vaciar carrito YA — el pedido está guardado, el usuario no espera más ──
        $identifier = TempCart::getCartIdentifier();
        TempCart::where('status', 0)
            ->where($identifier['column'], $identifier['value'])
            ->update(['status' => 1, 'updated_at' => now()]);

        // ── 3. Llamada a Profit API DESPUÉS de que la respuesta HTTP sea enviada ──
        // app()->terminating() ejecuta el callback tras $response->send(), sin bloquear al usuario.
        $pedidoId = $pedido->id;
        $coCli    = $user->co_cli;
        $payload  = [
            'cod_pedido'  => $pedidoId,
            'cod_cliente' => $coCli,
            'items'       => $validItems->map(fn($item) => [
                'co_art'     => $item['co_art'],
                'precio'     => (float) ($item['prec_agr'] ?? 0),
                'quantitySC' => (int) ($item['options']['quantitySC'] ?? 0),
                'quantityBQ' => (int) ($item['options']['quantityBQ'] ?? 0),
            ])->values()->toArray(),
        ];
        $orderUrl = config('services.profit.order_url', 'https://apiweb.cristmedicals.com/api/pedidos/pedido-profit');

        app()->terminating(function () use ($payload, $pedidoId, $coCli, $orderUrl) {
            try {
                $response = Http::timeout(15)->connectTimeout(5)->post($orderUrl, $payload);

                if ($response->successful()) {
                    Pedido::find($pedidoId)?->update(['status' => 1]);
                    Pedido::where('cod_cliente', $coCli)->where('status', 2)->where('id', '!=', $pedidoId)->update(['status' => 3]);
                    Log::info("ProcessCart async: Pedido #{$pedidoId} enviado a Profit OK.");
                    return;
                }

                if ($response->status() === 409) {
                    $body = $response->json();
                    if (($body['resultado']['duplicado'] ?? false) && ($body['resultado']['pedido_montado'] ?? false)) {
                        Pedido::find($pedidoId)?->update(['status' => 1]);
                        Log::info("ProcessCart async: Pedido #{$pedidoId} ya existía en Profit — marcado enviado.");
                        return;
                    }
                }

                // Fallo no recuperable → pedido queda en status=2 (retenido), admin puede reenviar
                Log::error("ProcessCart async: Pedido #{$pedidoId} retenido.", [
                    'http_status' => $response->status(),
                    'body'        => $response->body(),
                ]);

            } catch (\Exception $e) {
                Log::error("ProcessCart async: excepción en pedido #{$pedidoId}: " . $e->getMessage());
            }
        });

        // Retornar true inmediatamente — el carrito ya fue vaciado y el pedido guardado
        return true;
    }

    private function _sendOrderToExternalEndpoint(
        string $endpointUrl,
        array $payload,
        int $maxRetries = 3,
        int $retryDelayMs = 500
    ): bool {
        $pedidoId      = $payload['cod_pedido'] ?? 'unknown';
        $attempt       = 0;
        $lastException = null;

        while ($attempt < $maxRetries) {
            $attempt++;

            try {
                $response = Http::timeout(10)
                    ->connectTimeout(5)
                    ->post($endpointUrl, $payload);

                if ($response->successful()) {
                    return true;
                }

                // ✅ 409 con duplicado confirmado = el pedido SÍ fue procesado
                if ($response->status() === 409) {
                    $body          = $response->json();
                    $esDuplicado   = $body['resultado']['duplicado'] ?? false;
                    $pedidoMontado = $body['resultado']['pedido_montado'] ?? false;

                    if ($esDuplicado && $pedidoMontado) {
                        Log::info("_sendOrderToExternalEndpoint: Pedido #{$pedidoId} ya estaba procesado externamente. Se trata como éxito.", [
                            'fact_nums' => $body['resultado']['fact_nums'] ?? [],
                        ]);
                        return true; // processCart limpiará el carrito y pondrá status 1
                    }
                }

                if ($response->clientError()) {
                    Log::error("_sendOrderToExternalEndpoint: Error permanente (4xx) para pedido #{$pedidoId}. No se reintentará.", [
                        'status'   => $response->status(),
                        'response' => $response->body(),
                        'payload'  => $payload,
                    ]);
                    return false;
                }

            } catch (\Illuminate\Http\Client\ConnectionException $e) {
                $lastException = $e;

            } catch (Exception $e) {
                Log::error("_sendOrderToExternalEndpoint: Error inesperado para pedido #{$pedidoId}. Se abortan los reintentos.", [
                    'message' => $e->getMessage(),
                    'payload' => $payload,
                ]);
                throw $e;
            }

            if ($attempt < $maxRetries) {
                $sleepMs = $retryDelayMs * $attempt;
                usleep($sleepMs * 1000);
            }
        }

        Log::error("_sendOrderToExternalEndpoint: Se agotaron {$maxRetries} intentos para pedido #{$pedidoId}.", [
            'ultimo_error' => $lastException?->getMessage(),
            'payload'      => $payload,
        ]);

        if ($lastException) {
            throw $lastException;
        }

        return false;
    }

}
